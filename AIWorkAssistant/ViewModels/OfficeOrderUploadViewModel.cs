using System.Collections.ObjectModel;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using AIWorkAssistant.Models.HkOrder;
using AIWorkAssistant.Services.HkOrder;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class OfficeOrderUploadViewModel : ObservableObject
{
    [ObservableProperty] private string _orderFolderPath = string.Empty;
    [ObservableProperty] private string _targetUrl = string.Empty;
    [ObservableProperty] private string _orderPageUrl = string.Empty;
    [ObservableProperty] private string _username = string.Empty;
    [ObservableProperty] private string _password = string.Empty;
    [ObservableProperty] private string _apiBaseUrl = string.Empty;
    [ObservableProperty] private string _apiKey = string.Empty;
    [ObservableProperty] private string _modelName = string.Empty;
    [ObservableProperty] private string _saleType = string.Empty;
    [ObservableProperty] private string _moneyType = "人民币";
    [ObservableProperty] private string _rate = "1";
    [ObservableProperty] private string _marketType = "老市场";
    [ObservableProperty] private string _browserType = "Chromium";
    [ObservableProperty] private bool _headless;
    [ObservableProperty] private int _slowMo;
    [ObservableProperty] private int _defaultTimeout = 30000;

    [ObservableProperty] private bool _isRunning;
    [ObservableProperty] private int _progressValue;
    [ObservableProperty] private string _progressText = string.Empty;
    [ObservableProperty] private string _statusMessage = string.Empty;

    [ObservableProperty] private bool _isCaptchaVisible;
    [ObservableProperty] private string _captchaImagePath = string.Empty;
    [ObservableProperty] private string _captchaInput = string.Empty;

    [ObservableProperty] private bool _isConfirmVisible;
    [ObservableProperty] private string _previewImagePath = string.Empty;

    public ObservableCollection<AgentLogEntry> Logs { get; } = new();

    private HkOrderWorkflowService? _workflow;
    private CancellationTokenSource? _cts;
    private TaskCompletionSource<string?>? _captchaTcs;
    private TaskCompletionSource<bool>? _confirmTcs;

    public async Task LoadAsync()
    {
        await using var db = new AppDbContext();
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);

        OrderFolderPath = settings.GetValueOrDefault("HkOrder_OrderFolderPath", "");
        TargetUrl = settings.GetValueOrDefault("HkOrder_TargetUrl", settings.GetValueOrDefault("SystemUrl", ""));
        OrderPageUrl = settings.GetValueOrDefault("HkOrder_OrderPageUrl", "");
        Username = settings.GetValueOrDefault("HkOrder_Username", settings.GetValueOrDefault("SystemUsername", ""));
        Password = settings.GetValueOrDefault("HkOrder_Password", settings.GetValueOrDefault("SystemPassword", ""));
        ApiBaseUrl = settings.GetValueOrDefault("HkOrder_ApiBaseUrl", settings.GetValueOrDefault("AiBaseUrl", ""));
        ApiKey = settings.GetValueOrDefault("HkOrder_ApiKey", settings.GetValueOrDefault("AiApiKey", ""));
        ModelName = settings.GetValueOrDefault("HkOrder_ModelName", settings.GetValueOrDefault("AiModel", ""));
        SaleType = settings.GetValueOrDefault("HkOrder_SaleType", settings.GetValueOrDefault("DefaultSaleType", "煤矿"));
        MoneyType = settings.GetValueOrDefault("HkOrder_MoneyType", settings.GetValueOrDefault("DefaultMoneyType", "人民币"));
        Rate = settings.GetValueOrDefault("HkOrder_Rate", settings.GetValueOrDefault("DefaultRate", "1"));
        MarketType = settings.GetValueOrDefault("HkOrder_MarketType", settings.GetValueOrDefault("DefaultMarketType", "老市场"));
        BrowserType = settings.GetValueOrDefault("HkOrder_BrowserType", "Chromium");
        Headless = bool.TryParse(settings.GetValueOrDefault("HkOrder_Headless", "false"), out var headless) && headless;
        SlowMo = int.TryParse(settings.GetValueOrDefault("HkOrder_SlowMo", "0"), out var slowMo) ? slowMo : 0;
        DefaultTimeout = int.TryParse(settings.GetValueOrDefault("HkOrder_DefaultTimeout", "30000"), out var timeout) ? timeout : 30000;
    }

    [RelayCommand]
    private async Task SaveSettingsAsync()
    {
        await SaveSettingsCoreAsync();
        StatusMessage = "✅ 办公订单上传配置已保存";
    }

    [RelayCommand]
    private async Task StartAsync()
    {
        if (IsRunning) return;

        if (string.IsNullOrWhiteSpace(TargetUrl))
        {
            AddLog("请先填写登录网址", true);
            return;
        }

        if (string.IsNullOrWhiteSpace(OrderPageUrl))
        {
            AddLog("请先填写新增订单页面地址", true);
            return;
        }

        if (string.IsNullOrWhiteSpace(OrderFolderPath) || !Directory.Exists(OrderFolderPath))
        {
            AddLog("请先填写有效的订单文件夹", true);
            return;
        }

        var orderFiles = OrderFileReaderService.GetOrderFiles(OrderFolderPath);
        if (orderFiles.Length == 0)
        {
            AddLog("所选文件夹中没有找到 .doc/.docx/.xls/.xlsx 订单文件", true);
            return;
        }

        await SaveSettingsCoreAsync();
        Logs.Clear();
        IsRunning = true;
        ProgressValue = 0;
        ProgressText = "启动中...";
        StatusMessage = string.Empty;
        _cts = new CancellationTokenSource();

        try
        {
            AddLog($"共 {orderFiles.Length} 个订单文件，开始处理...", false);
            _workflow = new HkOrderWorkflowService(
                BuildSettings(),
                message => Dispatcher.UIThread.Post(() => AddLog(message, false)),
                RequestCaptchaInputAsync,
                RequestSubmitConfirmationAsync);

            var progress = new Progress<(int Percent, string Text)>(p =>
            {
                ProgressValue = p.Percent;
                ProgressText = p.Text;
            });

            await _workflow.RunAsync(orderFiles, progress, _cts.Token);
            StatusMessage = "✅ 全部订单处理完成";
        }
        catch (OperationCanceledException)
        {
            AddLog("任务已取消", true);
            StatusMessage = "任务已取消";
        }
        catch (Exception ex)
        {
            AddLog($"[错误] {ex.Message}", true);
            StatusMessage = $"❌ {ex.Message}";
        }
        finally
        {
            if (_workflow != null)
            {
                await _workflow.DisposeAsync();
                _workflow = null;
            }

            _cts?.Dispose();
            _cts = null;
            IsRunning = false;
        }
    }

    [RelayCommand]
    private void Stop()
    {
        _cts?.Cancel();
        _captchaTcs?.TrySetCanceled();
        _confirmTcs?.TrySetResult(false);
        IsCaptchaVisible = false;
        IsConfirmVisible = false;
        AddLog("正在停止...", false);
    }

    [RelayCommand]
    private void SubmitCaptcha()
    {
        if (_captchaTcs == null) return;
        var input = CaptchaInput.Trim();
        if (string.IsNullOrWhiteSpace(input)) return;

        _captchaTcs.TrySetResult(input);
        CaptchaInput = string.Empty;
        IsCaptchaVisible = false;
    }

    [RelayCommand]
    private void CancelCaptcha()
    {
        _captchaTcs?.TrySetResult(null);
        CaptchaInput = string.Empty;
        IsCaptchaVisible = false;
    }

    [RelayCommand]
    private void ConfirmSubmit()
    {
        _confirmTcs?.TrySetResult(true);
        IsConfirmVisible = false;
    }

    [RelayCommand]
    private void SkipSubmit()
    {
        _confirmTcs?.TrySetResult(false);
        IsConfirmVisible = false;
    }

    private Task<string?> RequestCaptchaInputAsync(string captchaPath)
    {
        _captchaTcs = new TaskCompletionSource<string?>();
        Dispatcher.UIThread.Post(() =>
        {
            CaptchaImagePath = captchaPath;
            CaptchaInput = string.Empty;
            IsCaptchaVisible = true;
            AddLog("⚠️ 请输入验证码", false);
        });
        return _captchaTcs.Task;
    }

    private Task<bool> RequestSubmitConfirmationAsync(string screenshotPath)
    {
        _confirmTcs = new TaskCompletionSource<bool>();
        Dispatcher.UIThread.Post(() =>
        {
            PreviewImagePath = screenshotPath;
            IsConfirmVisible = true;
            AddLog("请检查订单预览截图，确认后提交", false);
        });
        return _confirmTcs.Task;
    }

    private HkOrderSettings BuildSettings() => new()
    {
        BrowserType = BrowserType,
        Headless = Headless,
        TargetUrl = TargetUrl,
        OrderPageUrl = OrderPageUrl,
        Username = Username,
        Password = Password,
        DefaultTimeout = DefaultTimeout,
        SlowMo = SlowMo,
        OrderFolderPath = OrderFolderPath,
        ApiKey = ApiKey,
        ApiBaseUrl = ApiBaseUrl,
        ModelName = ModelName,
        SaleType = SaleType,
        MoneyType = MoneyType,
        Rate = Rate,
        MarketType = MarketType
    };

    private async Task SaveSettingsCoreAsync()
    {
        await using var db = new AppDbContext();
        await UpsertSetting(db, "HkOrder_OrderFolderPath", OrderFolderPath);
        await UpsertSetting(db, "HkOrder_TargetUrl", TargetUrl);
        await UpsertSetting(db, "HkOrder_OrderPageUrl", OrderPageUrl);
        await UpsertSetting(db, "HkOrder_Username", Username);
        await UpsertSetting(db, "HkOrder_Password", Password);
        await UpsertSetting(db, "HkOrder_ApiBaseUrl", ApiBaseUrl);
        await UpsertSetting(db, "HkOrder_ApiKey", ApiKey);
        await UpsertSetting(db, "HkOrder_ModelName", ModelName);
        await UpsertSetting(db, "HkOrder_SaleType", SaleType);
        await UpsertSetting(db, "HkOrder_MoneyType", MoneyType);
        await UpsertSetting(db, "HkOrder_Rate", Rate);
        await UpsertSetting(db, "HkOrder_MarketType", MarketType);
        await UpsertSetting(db, "HkOrder_BrowserType", BrowserType);
        await UpsertSetting(db, "HkOrder_Headless", Headless.ToString());
        await UpsertSetting(db, "HkOrder_SlowMo", SlowMo.ToString());
        await UpsertSetting(db, "HkOrder_DefaultTimeout", DefaultTimeout.ToString());
        await db.SaveChangesAsync();
    }

    private static async Task UpsertSetting(AppDbContext db, string key, string value)
    {
        var existing = await db.AppSettings.FindAsync(key);
        if (existing != null)
            existing.Value = value ?? string.Empty;
        else
            db.AppSettings.Add(new AppSetting { Key = key, Value = value ?? string.Empty });
    }

    private void AddLog(string message, bool isError)
    {
        Logs.Add(new AgentLogEntry
        {
            Timestamp = DateTime.Now,
            Message = message,
            IsError = isError
        });
    }
}

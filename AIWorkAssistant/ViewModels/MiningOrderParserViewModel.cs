using System.Text.Encodings.Web;
using System.Text.Json;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using AIWorkAssistant.Models.HkOrder;
using AIWorkAssistant.Services.HkOrder;
using AIWorkAssistant.Services.OrderParsing;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class MiningOrderParserViewModel : ObservableObject
{
    private static readonly JsonSerializerOptions OutputJsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    [ObservableProperty] private string _filePath = string.Empty;
    [ObservableProperty] private string _apiBaseUrl = "https://yxai.chat/v1";
    [ObservableProperty] private string _modelName = "glm-5.1";
    [ObservableProperty] private string _apiFormat = "claude_messages";
    [ObservableProperty] private string _apiKey = string.Empty;
    [ObservableProperty] private bool _useAi = true;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _targetUrl = string.Empty;
    [ObservableProperty] private string _orderPageUrl = string.Empty;
    [ObservableProperty] private string _username = string.Empty;
    [ObservableProperty] private string _password = string.Empty;
    [ObservableProperty] private string _saleType = "煤矿";
    [ObservableProperty] private string _moneyType = "人民币";
    [ObservableProperty] private string _rate = "1";
    [ObservableProperty] private string _marketType = "老市场";
    [ObservableProperty] private string _statusMessage = "请选择订单 Word/TXT 文件";
    [ObservableProperty] private string _resultJson = "等待解析...";
    [ObservableProperty] private string _uploadJson = "等待生成上传 JSON...";
    [ObservableProperty] private string _logText = string.Empty;
    [ObservableProperty] private bool _isCaptchaVisible;
    [ObservableProperty] private string _captchaImagePath = string.Empty;
    [ObservableProperty] private string _captchaInput = string.Empty;
    [ObservableProperty] private bool _isConfirmVisible;
    [ObservableProperty] private string _previewImagePath = string.Empty;

    public System.Collections.ObjectModel.ObservableCollection<AgentLogEntry> Logs { get; } = new();

    private ParsedOrder? _lastParsedOrder;
    private OrderData? _lastUploadOrder;
    private PlaywrightService? _playwright;
    private BrowserAgent? _browserAgent;
    private TaskCompletionSource<string?>? _captchaTcs;
    private TaskCompletionSource<bool>? _confirmTcs;

    public async Task LoadAsync()
    {
        await using var db = new AppDbContext();
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);

        ApiBaseUrl = NormalizeBaseUrl(settings.GetValueOrDefault("AiBaseUrl", settings.GetValueOrDefault("ApiBaseUrl", ApiBaseUrl)));
        ApiKey = settings.GetValueOrDefault("AiApiKey", settings.GetValueOrDefault("ApiKey", ""));
        ModelName = settings.GetValueOrDefault("AiModelName", settings.GetValueOrDefault("ModelName", ModelName));
        ApiFormat = settings.GetValueOrDefault("AiApiFormat", InferApiFormat(ModelName));
        TargetUrl = settings.GetValueOrDefault("HkOrder_TargetUrl", settings.GetValueOrDefault("SystemUrl", ""));
        OrderPageUrl = settings.GetValueOrDefault("HkOrder_OrderPageUrl", "");
        Username = settings.GetValueOrDefault("HkOrder_Username", settings.GetValueOrDefault("SystemUsername", ""));
        Password = settings.GetValueOrDefault("HkOrder_Password", settings.GetValueOrDefault("SystemPassword", ""));
        SaleType = settings.GetValueOrDefault("HkOrder_SaleType", settings.GetValueOrDefault("DefaultSaleType", "煤矿"));
        MoneyType = settings.GetValueOrDefault("HkOrder_MoneyType", settings.GetValueOrDefault("DefaultMoneyType", "人民币"));
        Rate = settings.GetValueOrDefault("HkOrder_Rate", settings.GetValueOrDefault("DefaultRate", "1"));
        MarketType = settings.GetValueOrDefault("HkOrder_MarketType", settings.GetValueOrDefault("DefaultMarketType", "老市场"));
    }

    [RelayCommand]
    private async Task ParseAsync()
    {
        if (string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
        {
            StatusMessage = "请先选择有效的订单文件";
            return;
        }

        IsBusy = true;
        StatusMessage = "正在读取并解析订单...";
        ResultJson = "解析中...";

        try
        {
            var text = await DocumentTextExtractor.ExtractAsync(FilePath);
            var order = OrderParser.Parse(text);

            if (UseAi)
            {
                var key = ApiKey;
                if (string.IsNullOrWhiteSpace(key)) key = Environment.GetEnvironmentVariable("YXAI_API_KEY") ?? "";

                if (string.IsNullOrWhiteSpace(key))
                {
                    order.Validation.Warnings.Add("AI 已开启，但未配置 API Key。本次只返回规则解析结果。请在本页填写 API Key 或设置环境变量 YXAI_API_KEY。");
                }
                else
                {
                    var format = string.IsNullOrWhiteSpace(ApiFormat) ? InferApiFormat(ModelName) : ApiFormat;
                    order.AiReview = await YxAiClient.FallbackAndReviewAsync(
                        NormalizeBaseUrl(ApiBaseUrl),
                        key,
                        string.IsNullOrWhiteSpace(ModelName) ? "glm-5.1" : ModelName,
                        format,
                        text,
                        order);
                    order.Validation = OrderParser.Validate(order);
                }
            }

            _lastParsedOrder = order;
            _lastUploadOrder = MiningOrderUploadMapper.ToUploadOrder(order);
            ResultJson = JsonSerializer.Serialize(order, OutputJsonOptions);
            UploadJson = JsonSerializer.Serialize(_lastUploadOrder, OutputJsonOptions);
            StatusMessage = "✅ 解析完成，已生成上传 JSON";
        }
        catch (Exception ex)
        {
            StatusMessage = "❌ 解析失败";
            ResultJson = ex.ToString();
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void CopyResult()
    {
        StatusMessage = "请在右侧结果框中全选复制 JSON";
    }

    [RelayCommand]
    private async Task SaveResultAsync()
    {
        if (string.IsNullOrWhiteSpace(FilePath) || string.IsNullOrWhiteSpace(ResultJson) || ResultJson == "等待解析...")
        {
            StatusMessage = "没有可保存的解析结果";
            return;
        }

        var parsedPath = Path.ChangeExtension(FilePath, ".parsed.json");
        await File.WriteAllTextAsync(parsedPath, ResultJson);

        if (!string.IsNullOrWhiteSpace(UploadJson) && UploadJson != "等待生成上传 JSON...")
        {
            var uploadPath = Path.ChangeExtension(FilePath, ".upload.json");
            await File.WriteAllTextAsync(uploadPath, UploadJson);
            StatusMessage = $"已保存：{parsedPath} 和 {uploadPath}";
        }
        else
        {
            StatusMessage = $"已保存：{parsedPath}";
        }
    }

    [RelayCommand]
    private async Task UploadAsync()
    {
        if (_lastUploadOrder == null)
        {
            await ParseAsync();
            if (_lastUploadOrder == null) return;
        }

        if (string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
        {
            StatusMessage = "请先选择有效的订单文件";
            return;
        }

        if (string.IsNullOrWhiteSpace(TargetUrl) || string.IsNullOrWhiteSpace(OrderPageUrl) ||
            string.IsNullOrWhiteSpace(Username) || string.IsNullOrWhiteSpace(Password))
        {
            StatusMessage = "请先填写订单系统地址、新增订单页面、账号和密码";
            return;
        }

        IsBusy = true;
        StatusMessage = "正在通过浏览器上传订单...";

        try
        {
            await SaveUploadSettingsAsync();
            var agent = await EnsureBrowserAgentAsync();
            AddLog("开始填写订单系统表单...", false);
            await agent.NavigateToAddOrderAsync();
            await agent.FillFormAsync(_lastUploadOrder);
            await agent.UploadDocFileAsync(FilePath);

            var screenshotPath = await agent.TakeScreenshotAsync(Path.GetFileNameWithoutExtension(FilePath));
            PreviewImagePath = screenshotPath;
            AddLog($"已生成提交前截图：{screenshotPath}", false);

            var confirmed = await RequestSubmitConfirmationAsync(screenshotPath);
            if (confirmed)
            {
                await agent.SubmitAsync();
                AddLog("订单已提交。", false);
                StatusMessage = "✅ 上传并提交完成";
                await CloseBrowserAsync();
            }
            else
            {
                AddLog("用户选择跳过提交，浏览器保持当前页面便于检查。", false);
                StatusMessage = "已跳过提交，浏览器保持打开";
            }
        }
        catch (Exception ex)
        {
            AddLog(ex.Message, true);
            StatusMessage = "❌ 上传失败";
        }
        finally
        {
            IsBusy = false;
        }
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

    private async Task<BrowserAgent> EnsureBrowserAgentAsync()
    {
        if (_browserAgent != null && _playwright?.Page != null && _playwright.IsRunning)
        {
            return _browserAgent;
        }

        _playwright = new PlaywrightService();
        var settings = BuildUploadSettings();
        AddLog("正在启动浏览器并登录订单系统...", false);
        await _playwright.StartAsync(settings, message => AddLog(message, false));
        if (_playwright.Page == null) throw new InvalidOperationException("浏览器页面启动失败，无法上传订单。");
        _browserAgent = new BrowserAgent(_playwright.Page, settings, message => AddLog(message, false));
        await _browserAgent.LoginAsync(RequestCaptchaInputAsync);
        AddLog("登录完成。", false);
        return _browserAgent;
    }

    private HkOrderSettings BuildUploadSettings() => new()
    {
        BrowserType = "Chromium",
        Headless = false,
        TargetUrl = TargetUrl,
        OrderPageUrl = OrderPageUrl,
        Username = Username,
        Password = Password,
        DefaultTimeout = 30000,
        SlowMo = 0,
        SaleType = SaleType,
        MoneyType = MoneyType,
        Rate = Rate,
        MarketType = MarketType
    };

    private async Task SaveUploadSettingsAsync()
    {
        await using var db = new AppDbContext();
        await UpsertSetting(db, "HkOrder_TargetUrl", TargetUrl);
        await UpsertSetting(db, "HkOrder_OrderPageUrl", OrderPageUrl);
        await UpsertSetting(db, "HkOrder_Username", Username);
        await UpsertSetting(db, "HkOrder_Password", Password);
        await UpsertSetting(db, "HkOrder_SaleType", SaleType);
        await UpsertSetting(db, "HkOrder_MoneyType", MoneyType);
        await UpsertSetting(db, "HkOrder_Rate", Rate);
        await UpsertSetting(db, "HkOrder_MarketType", MarketType);
        await db.SaveChangesAsync();
    }

    private static async Task UpsertSetting(AppDbContext db, string key, string value)
    {
        var existing = await db.AppSettings.FindAsync(key);
        if (existing != null) existing.Value = value ?? string.Empty;
        else db.AppSettings.Add(new AppSetting { Key = key, Value = value ?? string.Empty });
    }

    private Task<string?> RequestCaptchaInputAsync(string captchaPath)
    {
        _captchaTcs = new TaskCompletionSource<string?>();
        Dispatcher.UIThread.Post(() =>
        {
            CaptchaImagePath = captchaPath;
            CaptchaInput = string.Empty;
            IsCaptchaVisible = true;
            AddLog("请输入验证码", false);
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

    private async Task CloseBrowserAsync()
    {
        if (_playwright != null)
        {
            await _playwright.DisposeAsync();
            _playwright = null;
        }
        _browserAgent = null;
    }

    private void AddLog(string message, bool isError)
    {
        Logs.Add(new AgentLogEntry
        {
            Timestamp = DateTime.Now,
            Message = message,
            IsError = isError
        });
        LogText = string.Join(Environment.NewLine, Logs.Select(log => log.Display));
    }

    private static string InferApiFormat(string model)
    {
        return model.StartsWith("glm-", StringComparison.OrdinalIgnoreCase)
            ? "claude_messages"
            : "openai_chat";
    }

    private static string NormalizeBaseUrl(string value)
    {
        var url = string.IsNullOrWhiteSpace(value) ? "https://yxai.chat/v1" : value.Trim().TrimEnd('/');
        return url.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) ? url : url + "/v1";
    }
}

using System.Collections.ObjectModel;
using System.Text.Encodings.Web;
using System.Text.Json;
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
    private const string DefaultAiBaseUrl = "https://yxai.chat";
    private const string DefaultAiModel = "glm-5.1";

    private static readonly JsonSerializerOptions ParsedJsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    [ObservableProperty] private string _orderFolderPath = string.Empty;
    [ObservableProperty] private string _targetUrl = string.Empty;
    [ObservableProperty] private string _orderPageUrl = string.Empty;
    [ObservableProperty] private string _username = string.Empty;
    [ObservableProperty] private string _password = string.Empty;
    [ObservableProperty] private string _saleType = string.Empty;
    [ObservableProperty] private string _moneyType = "人民币";
    [ObservableProperty] private string _rate = "1";
    [ObservableProperty] private string _marketType = "老市场";
    [ObservableProperty] private string _browserType = "Chromium";
    [ObservableProperty] private bool _headless;
    [ObservableProperty] private int _slowMo;
    [ObservableProperty] private int _defaultTimeout = 30000;
    [ObservableProperty] private int _scanIntervalMinutes = 10;

    [ObservableProperty] private bool _isRunning;
    [ObservableProperty] private int _progressValue;
    [ObservableProperty] private string _progressText = string.Empty;
    [ObservableProperty] private string _statusMessage = string.Empty;
    [ObservableProperty] private string _lastScanTime = string.Empty;
    [ObservableProperty] private string _nextScanTime = string.Empty;
    [ObservableProperty] private string _logText = string.Empty;
    [ObservableProperty] private int _parsedFileCount;

    [ObservableProperty] private bool _isCaptchaVisible;
    [ObservableProperty] private string _captchaImagePath = string.Empty;
    [ObservableProperty] private string _captchaInput = string.Empty;

    [ObservableProperty] private bool _isConfirmVisible;
    [ObservableProperty] private string _previewImagePath = string.Empty;

    public ObservableCollection<AgentLogEntry> Logs { get; } = new();

    private HkOrderWorkflowService? _workflow;
    private PlaywrightService? _playwright;
    private BrowserAgent? _browserAgent;
    private CancellationTokenSource? _cts;
    private TaskCompletionSource<string?>? _captchaTcs;
    private TaskCompletionSource<bool>? _confirmTcs;
    private readonly HashSet<string> _processedFileKeys = new(StringComparer.OrdinalIgnoreCase);

    public async Task LoadAsync()
    {
        await using var db = new AppDbContext();
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);

        OrderFolderPath = settings.GetValueOrDefault("HkOrder_OrderFolderPath", "");
        TargetUrl = settings.GetValueOrDefault("HkOrder_TargetUrl", settings.GetValueOrDefault("SystemUrl", ""));
        OrderPageUrl = settings.GetValueOrDefault("HkOrder_OrderPageUrl", "");
        Username = settings.GetValueOrDefault("HkOrder_Username", settings.GetValueOrDefault("SystemUsername", ""));
        Password = settings.GetValueOrDefault("HkOrder_Password", settings.GetValueOrDefault("SystemPassword", ""));
        SaleType = settings.GetValueOrDefault("HkOrder_SaleType", settings.GetValueOrDefault("DefaultSaleType", "煤矿"));
        MoneyType = settings.GetValueOrDefault("HkOrder_MoneyType", settings.GetValueOrDefault("DefaultMoneyType", "人民币"));
        Rate = settings.GetValueOrDefault("HkOrder_Rate", settings.GetValueOrDefault("DefaultRate", "1"));
        MarketType = settings.GetValueOrDefault("HkOrder_MarketType", settings.GetValueOrDefault("DefaultMarketType", "老市场"));
        BrowserType = settings.GetValueOrDefault("HkOrder_BrowserType", "Chromium");
        Headless = bool.TryParse(settings.GetValueOrDefault("HkOrder_Headless", "false"), out var headless) && headless;
        SlowMo = int.TryParse(settings.GetValueOrDefault("HkOrder_SlowMo", "0"), out var slowMo) ? slowMo : 0;
        DefaultTimeout = int.TryParse(settings.GetValueOrDefault("HkOrder_DefaultTimeout", "30000"), out var timeout) ? timeout : 30000;
        ScanIntervalMinutes = int.TryParse(settings.GetValueOrDefault("HkOrder_ScanIntervalMinutes", "10"), out var interval)
            ? Math.Max(1, interval)
            : 10;
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

        if (string.IsNullOrWhiteSpace(OrderFolderPath) || !Directory.Exists(OrderFolderPath))
        {
            AddLog("请先填写有效的订单文件夹", true);
            return;
        }

        if (string.IsNullOrWhiteSpace(TargetUrl) ||
            string.IsNullOrWhiteSpace(OrderPageUrl) ||
            string.IsNullOrWhiteSpace(Username) ||
            string.IsNullOrWhiteSpace(Password))
        {
            AddLog("请先填写系统登录网址、新增订单页面地址、用户名和密码", true);
            return;
        }

        var aiSettings = await LoadGlobalAiSettingsAsync();
        if (string.IsNullOrWhiteSpace(aiSettings.ApiBaseUrl) ||
            string.IsNullOrWhiteSpace(aiSettings.ApiKey) ||
            string.IsNullOrWhiteSpace(aiSettings.ModelName))
        {
            AddLog("请先到“设置”里填写 AI 地址、API Key 和模型名称", true);
            return;
        }

        await SaveSettingsCoreAsync();
        Logs.Clear();
        LogText = string.Empty;
        IsRunning = true;
        ProgressValue = 0;
        ProgressText = "监听启动中...";
        StatusMessage = string.Empty;
        LastScanTime = string.Empty;
        NextScanTime = string.Empty;
        _cts = new CancellationTokenSource();

        try
        {
            AddLog($"开始监听：{OrderFolderPath}", false);
            AddLog($"扫描间隔：{ScanIntervalMinutes} 分钟。每次扫描会解析新增或修改过的 Word 文档并输出 .ai.json。", false);
            await RunFolderMonitorAsync(_cts.Token);
        }
        catch (OperationCanceledException)
        {
            AddLog("监听已停止", false);
            if (string.IsNullOrWhiteSpace(StatusMessage))
            {
                StatusMessage = "监听已停止";
            }
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

            if (_playwright != null)
            {
                await _playwright.DisposeAsync();
                _playwright = null;
            }

            _browserAgent = null;

            _cts?.Dispose();
            _cts = null;
            IsRunning = false;
            NextScanTime = string.Empty;
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

    private async Task RunFolderMonitorAsync(CancellationToken ct)
    {
        while (true)
        {
            ct.ThrowIfCancellationRequested();

            await ScanFolderOnceAsync(ct);

            var delay = TimeSpan.FromMinutes(Math.Max(1, ScanIntervalMinutes));
            var next = DateTime.Now.Add(delay);
            NextScanTime = next.ToString("yyyy-MM-dd HH:mm:ss");
            ProgressText = $"监听中，下次扫描：{NextScanTime}";
            await Task.Delay(delay, ct);
        }
    }

    private async Task ScanFolderOnceAsync(CancellationToken ct)
    {
        LastScanTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        NextScanTime = string.Empty;
        ProgressValue = 0;
        ProgressText = "正在扫描订单文件夹...";

        var orderFiles = Directory
            .EnumerateFiles(OrderFolderPath, "*.*", SearchOption.TopDirectoryOnly)
            .Where(IsWordOrderFile)
            .OrderBy(file => file)
            .ToArray();

        if (orderFiles.Length == 0)
        {
            AddLog($"[{LastScanTime}] 未找到 .doc/.docx 订单文件", false);
            ProgressText = "未找到 Word 订单文件";
            return;
        }

        var filesToParse = orderFiles
            .Where(file => !_processedFileKeys.Contains(CreateFileKey(file)))
            .ToArray();

        if (filesToParse.Length == 0)
        {
            AddLog($"[{LastScanTime}] 扫描到 {orderFiles.Length} 个 Word 文件，没有新增或修改", false);
            ProgressValue = 100;
            ProgressText = "本轮没有新增或修改文件";
            return;
        }

        AddLog($"[{LastScanTime}] 扫描到 {orderFiles.Length} 个 Word 文件，待解析 {filesToParse.Length} 个", false);

        for (var i = 0; i < filesToParse.Length; i++)
        {
            ct.ThrowIfCancellationRequested();

            var file = filesToParse[i];
            var fileName = Path.GetFileName(file);
            ProgressValue = (int)((i + 1) / (double)filesToParse.Length * 100);
            ProgressText = $"解析 {i + 1}/{filesToParse.Length}: {fileName}";

            try
            {
                var parseResult = await ParseOrderFileToJsonAsync(file, ct);
                await UploadParsedOrderAsync(file, parseResult.Data, ct);
                _processedFileKeys.Add(CreateFileKey(file));
                ParsedFileCount++;
            }
            catch (OrderItemCatalogMatchException ex)
            {
                AddLog($"[物品匹配错误] {fileName}", true);
                foreach (var line in ex.Message.Split(Environment.NewLine))
                {
                    AddLog(line, true);
                }

                StatusMessage = "物品匹配失败，已停止解析";
                ProgressText = "物品匹配失败，已停止解析";
                _cts?.Cancel();
                return;
            }
            catch (OrderSelectionNotFoundException ex)
            {
                AddLog($"[选择项错误] {fileName}", true);
                foreach (var line in ex.Message.Split(Environment.NewLine))
                {
                    AddLog(line, true);
                }

                await CloseBrowserAfterSelectionErrorAsync();
                StatusMessage = "选择项不存在，已退出浏览器并停止监听";
                ProgressText = "选择项不存在，已退出浏览器并停止监听";
                _cts?.Cancel();
                return;
            }
            catch (Exception ex)
            {
                AddLog($"[错误] {fileName}: {ex.Message}", true);
            }
        }

        StatusMessage = $"✅ 本轮扫描完成，累计解析 {ParsedFileCount} 个文件";
        ProgressText = "本轮扫描完成";
    }

    private async Task<AiParseResult> ParseOrderFileToJsonAsync(string file, CancellationToken ct)
    {
        var fileName = Path.GetFileName(file);
        AddLog($"读取 Word：{fileName}", false);

        var docText = await Task.Run(() => OrderFileReaderService.ReadText(file), ct);
        var docTextPath = file + ".txt";
        await File.WriteAllTextAsync(docTextPath, docText, ct);

        AddLog($"  文档长度：{docText.Length} 字符", false);
        AddLog($"  已输出提取文本：{docTextPath}", false);
        if (docText.Length > 0)
        {
            AddLog($"  前200字：{docText[..Math.Min(200, docText.Length)]}", false);
        }

        AddLog("  AI 解析中...", false);
        var aiSettings = await LoadGlobalAiSettingsAsync();
        var ai = new AiService(aiSettings.ApiKey, aiSettings.ApiBaseUrl, aiSettings.ModelName);
        var parseResult = await ai.ParseOrderAsync(docText, ct, message => AddLog(message, false));

        await WriteParseArtifactsAsync(file, parseResult, ct);
        AddLog($"  解析完成：使用单位 {parseResult.Data.CustomerName}，产品及配件 {parseResult.Data.Items.Count} 行", false);
        return parseResult;
    }

    private async Task UploadParsedOrderAsync(string file, OrderData orderData, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        var fileName = Path.GetFileName(file);
        var agent = await EnsureBrowserAgentAsync(ct);

        AddLog($"  开始上传到网站：{fileName}", false);
        await agent.NavigateToAddOrderAsync();
        await agent.FillFormAsync(orderData);
        await agent.UploadDocFileAsync(file);

        var screenshotPath = await agent.TakeScreenshotAsync(Path.GetFileNameWithoutExtension(fileName));
        AddLog($"  已生成提交前截图：{screenshotPath}", false);

        var confirmed = await RequestSubmitConfirmationAsync(screenshotPath);
        ct.ThrowIfCancellationRequested();

        if (!confirmed)
        {
            AddLog($"  已跳过提交：{fileName}", false);
            return;
        }

        await agent.SubmitAsync();
        AddLog($"  已提交：{fileName}", false);
        await CloseBrowserAfterSubmitSuccessAsync();
    }

    private async Task<BrowserAgent> EnsureBrowserAgentAsync(CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        if (_browserAgent != null && _playwright?.Page != null && _playwright.IsRunning)
        {
            return _browserAgent;
        }

        _playwright = new PlaywrightService();
        var settings = BuildSettings();

        AddLog("  正在启动浏览器并登录系统...", false);
        await _playwright.StartAsync(settings, message => AddLog(message, false));

        if (_playwright.Page == null)
        {
            throw new InvalidOperationException("浏览器页面启动失败，无法上传订单。");
        }

        _browserAgent = new BrowserAgent(_playwright.Page, settings, message => AddLog(message, false));
        await _browserAgent.LoginAsync(RequestCaptchaInputAsync);
        AddLog("  系统登录完成，后续订单将复用当前浏览器会话。", false);
        return _browserAgent;
    }

    private async Task CloseBrowserAfterSelectionErrorAsync()
    {
        AddLog("  找不到需要选择的项，正在退出浏览器...", true);
        if (_playwright != null)
        {
            await _playwright.DisposeAsync();
            _playwright = null;
        }

        _browserAgent = null;
        AddLog("  浏览器已退出。", true);
    }

    private async Task CloseBrowserAfterSubmitSuccessAsync()
    {
        AddLog("  提交成功，正在退出浏览器...", false);
        if (_playwright != null)
        {
            await _playwright.DisposeAsync();
            _playwright = null;
        }

        _browserAgent = null;
        AddLog("  浏览器已退出。", false);
    }

    private async Task WriteParseArtifactsAsync(string file, AiParseResult parseResult, CancellationToken ct)
    {
        var rawResponsePath = Path.ChangeExtension(file, ".ai.raw.json");
        var rawTextPath = Path.ChangeExtension(file, ".ai.raw.txt");
        var cleanedJsonPath = Path.ChangeExtension(file, ".ai.cleaned.json");
        var parsedJsonPath = Path.ChangeExtension(file, ".ai.json");

        await File.WriteAllTextAsync(rawResponsePath, parseResult.RawResponseJson, ct);
        await File.WriteAllTextAsync(rawTextPath, parseResult.RawModelText, ct);
        await File.WriteAllTextAsync(cleanedJsonPath, parseResult.CleanedJson, ct);

        var parsedJson = JsonSerializer.Serialize(parseResult.Data, ParsedJsonOptions);
        await File.WriteAllTextAsync(parsedJsonPath, parsedJson, ct);

        AddLog($"  已输出最终 JSON：{parsedJsonPath}", false);
        AddLog("  关键字段：", false);
        AddLog($"    使用单位：{parseResult.Data.CustomerName}", false);
        AddLog($"    合同金额：{parseResult.Data.ContractAmount}", false);
        AddLog($"    柔性网：{parseResult.Data.FlexibleNet}", false);
        AddLog($"    规格：{parseResult.Data.Spec}", false);
        AddLog($"    整编用的钢丝绳：{parseResult.Data.BraidedSteelWireRope}", false);
        AddLog($"    井下运输最大尺寸长：{parseResult.Data.UndergroundTransportMaxLength}", false);
        AddLog($"    销售经理：{parseResult.Data.SaleManager}", false);
        AddLog($"    销售部门：{parseResult.Data.SaleDept}", false);
    }

    private static bool IsWordOrderFile(string filePath)
    {
        var ext = Path.GetExtension(filePath);
        return ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".docx", StringComparison.OrdinalIgnoreCase);
    }

    private static string CreateFileKey(string filePath)
    {
        var info = new FileInfo(filePath);
        return $"{info.FullName}|{info.Length}|{info.LastWriteTimeUtc.Ticks}";
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
        ApiKey = string.Empty,
        ApiBaseUrl = string.Empty,
        ModelName = string.Empty,
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
        await UpsertSetting(db, "HkOrder_SaleType", SaleType);
        await UpsertSetting(db, "HkOrder_MoneyType", MoneyType);
        await UpsertSetting(db, "HkOrder_Rate", Rate);
        await UpsertSetting(db, "HkOrder_MarketType", MarketType);
        await UpsertSetting(db, "HkOrder_BrowserType", BrowserType);
        await UpsertSetting(db, "HkOrder_Headless", Headless.ToString());
        await UpsertSetting(db, "HkOrder_SlowMo", SlowMo.ToString());
        await UpsertSetting(db, "HkOrder_DefaultTimeout", DefaultTimeout.ToString());
        await UpsertSetting(db, "HkOrder_ScanIntervalMinutes", Math.Max(1, ScanIntervalMinutes).ToString());
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

    private static string FirstNonEmpty(params string?[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;

    private static async Task<(string ApiBaseUrl, string ApiKey, string ModelName)> LoadGlobalAiSettingsAsync()
    {
        await using var db = new AppDbContext();
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);

        return (
            FirstNonEmpty(settings.GetValueOrDefault("AiBaseUrl"), DefaultAiBaseUrl),
            settings.GetValueOrDefault("AiApiKey", string.Empty),
            FirstNonEmpty(settings.GetValueOrDefault("AiModel"), DefaultAiModel));
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
}

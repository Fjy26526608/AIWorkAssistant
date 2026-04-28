using System.Text.Encodings.Web;
using System.Text.Json;
using AIWorkAssistant.Models.HkOrder;

namespace AIWorkAssistant.Services.HkOrder;

public sealed class HkOrderWorkflowService : IAsyncDisposable
{
    private static readonly JsonSerializerOptions ParsedJsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly HkOrderSettings _settings;
    private readonly Action<string> _log;
    private readonly Func<string, Task<string?>> _requestCaptchaInput;
    private readonly Func<string, Task<bool>> _requestSubmitConfirmation;
    private PlaywrightService? _playwright;

    public HkOrderWorkflowService(
        HkOrderSettings settings,
        Action<string> log,
        Func<string, Task<string?>> requestCaptchaInput,
        Func<string, Task<bool>> requestSubmitConfirmation)
    {
        _settings = settings;
        _log = log;
        _requestCaptchaInput = requestCaptchaInput;
        _requestSubmitConfirmation = requestSubmitConfirmation;
    }

    public async Task RunAsync(IReadOnlyList<string> orderFiles, IProgress<(int Percent, string Text)>? progress, CancellationToken ct)
    {
        var parsedOrders = new List<(string FilePath, OrderData Data)>();

        for (var i = 0; i < orderFiles.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            var file = orderFiles[i];
            var fileName = Path.GetFileName(file);
            var percent = (int)((i + 1) / (double)orderFiles.Count * 50);
            progress?.Report((percent, $"解析 {i + 1}/{orderFiles.Count}"));

            _log($"[{i + 1}/{orderFiles.Count}] 读取：{fileName}");
            var docText = await Task.Run(() => OrderFileReaderService.ReadText(file), ct);
            var docTextPath = file + ".txt";
            await File.WriteAllTextAsync(docTextPath, docText, ct);

            _log($"  文档长度：{docText.Length} 字符");
            _log($"  已输出文档提取文本：{docTextPath}");
            if (docText.Length > 0)
            {
                _log($"  前200字：{docText[..Math.Min(200, docText.Length)]}");
            }

            _log("  AI 解析中...");
            var ai = new AiService(_settings.ApiKey, _settings.ApiBaseUrl, _settings.ModelName);
            var parseResult = await ai.ParseOrderAsync(docText, ct, _log);
            var orderData = parseResult.Data;

            await WriteParseArtifactsAsync(file, orderData, parseResult, ct);
            _log($"  解析完成：客户 {orderData.CustomerName}，物品 {orderData.Items.Count} 行");
            parsedOrders.Add((file, orderData));
        }

        _log($"全部解析完成，共 {parsedOrders.Count} 条订单。正在启动浏览器...");
        _playwright = new PlaywrightService();
        await _playwright.StartAsync(_settings, _log);
        _log($"浏览器已启动 [{_settings.BrowserType}]，已导航至 {_settings.TargetUrl}");

        var agent = new BrowserAgent(_playwright.Page!, _settings, _log);
        await agent.LoginAsync(captchaPath => RequestCaptchaInput(captchaPath, ct).GetAwaiter().GetResult());

        for (var i = 0; i < parsedOrders.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            var (filePath, orderData) = parsedOrders[i];
            var fileName = Path.GetFileName(filePath);
            var percent = 50 + (int)((i + 1) / (double)parsedOrders.Count * 50);
            progress?.Report((percent, $"上传 {i + 1}/{parsedOrders.Count}"));

            _log($"[{i + 1}/{parsedOrders.Count}] 处理：{fileName}");
            await agent.NavigateToAddOrderAsync();
            await agent.FillFormAsync(orderData);
            await agent.UploadDocFileAsync(filePath);

            var finalScreenshot = await agent.TakeScreenshotAsync($"order_{i + 1}");
            _log($"  截图：{finalScreenshot}");

            var confirmed = await _requestSubmitConfirmation(finalScreenshot);
            if (confirmed)
            {
                await agent.SubmitAsync();
                _log($"  已提交：{fileName}");
            }
            else
            {
                _log($"  已跳过：{fileName}");
            }
        }

        progress?.Report((100, "完成"));
        _log("全部订单处理完成。");
    }

    private async Task<string?> RequestCaptchaInput(string captchaPath, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        return await _requestCaptchaInput(captchaPath);
    }

    private async Task WriteParseArtifactsAsync(string file, OrderData orderData, AiParseResult parseResult, CancellationToken ct)
    {
        var rawResponsePath = Path.ChangeExtension(file, ".ai.raw.json");
        var rawTextPath = Path.ChangeExtension(file, ".ai.raw.txt");
        var cleanedJsonPath = Path.ChangeExtension(file, ".ai.cleaned.json");
        var parsedJsonPath = Path.ChangeExtension(file, ".ai.json");

        await File.WriteAllTextAsync(rawResponsePath, parseResult.RawResponseJson, ct);
        await File.WriteAllTextAsync(rawTextPath, parseResult.RawModelText, ct);
        await File.WriteAllTextAsync(cleanedJsonPath, parseResult.CleanedJson, ct);

        var parsedJson = JsonSerializer.Serialize(orderData, ParsedJsonOptions);
        await File.WriteAllTextAsync(parsedJsonPath, parsedJson, ct);

        _log($"  已输出 AI 原始响应：{rawResponsePath}");
        _log($"  已输出 AI 文本内容：{rawTextPath}");
        _log($"  已输出 AI 清洗JSON：{cleanedJsonPath}");
        _log($"  已输出 AI 最终结果：{parsedJsonPath}");
        _log("  AI解析结果(JSON)：");
        foreach (var line in parsedJson.Split('\n'))
        {
            _log($"    {line.TrimEnd('\r')}");
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_playwright != null)
        {
            await _playwright.DisposeAsync();
            _playwright = null;
        }
    }
}

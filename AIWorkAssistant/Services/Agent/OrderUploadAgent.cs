using System.IO;
using System.Text.Json;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.Services.Agent;

/// <summary>
/// 订单上传 Agent：协调文件监测、AI解析、浏览器自动化
/// </summary>
public class OrderUploadAgent : IDisposable
{
    private FileSystemWatcher? _watcher;
    private readonly Action<string, bool> _log; // message, isError
    private readonly Func<byte[], Task<string>> _requestCaptcha;
    private readonly HashSet<string> _processingFiles = new();
    private readonly object _lock = new();
    private bool _isRunning;

    public bool IsRunning => _isRunning;

    public OrderUploadAgent(Action<string, bool> log, Func<byte[], Task<string>> requestCaptcha)
    {
        _log = log;
        _requestCaptcha = requestCaptcha;
    }

    public void Start(string watchFolder)
    {
        if (_isRunning) return;

        if (!Directory.Exists(watchFolder))
            Directory.CreateDirectory(watchFolder);

        _watcher = new FileSystemWatcher(watchFolder)
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true
        };

        _watcher.Created += OnFileCreated;
        _watcher.Renamed += OnFileRenamed;
        _isRunning = true;

        _log("文件监测已启动: " + watchFolder, false);

        // 扫描已有文件
        var existing = Directory.GetFiles(watchFolder)
            .Where(f => IsWordFile(f)).ToList();

        if (existing.Count > 0)
        {
            _log($"发现 {existing.Count} 个待处理文件", false);
            foreach (var file in existing)
                _ = ProcessFileAsync(file);
        }
    }

    public void Stop()
    {
        if (_watcher != null)
        {
            _watcher.EnableRaisingEvents = false;
            _watcher.Created -= OnFileCreated;
            _watcher.Renamed -= OnFileRenamed;
        }
        _isRunning = false;
        _log("文件监测已停止", false);
    }

    private async void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        if (IsWordFile(e.FullPath))
        {
            await Task.Delay(1500); // 等待文件写入完成
            await ProcessFileAsync(e.FullPath);
        }
    }

    private async void OnFileRenamed(object sender, RenamedEventArgs e)
    {
        if (IsWordFile(e.FullPath))
        {
            await Task.Delay(1500);
            await ProcessFileAsync(e.FullPath);
        }
    }

    private async Task ProcessFileAsync(string filePath)
    {
        lock (_lock)
        {
            if (_processingFiles.Contains(filePath)) return;
            _processingFiles.Add(filePath);
        }

        var fileName = Path.GetFileName(filePath);
        var directory = Path.GetDirectoryName(filePath)!;
        var successDir = Path.Combine(directory, "已处理");
        var failDir = Path.Combine(directory, "失败");
        Directory.CreateDirectory(successDir);
        Directory.CreateDirectory(failDir);

        try
        {
            _log($"检测到文件: {fileName}", false);

            // 1. 读取文档
            _log("正在读取文档...", false);
            var text = WordReaderService.ReadDocument(filePath);
            if (string.IsNullOrWhiteSpace(text))
                throw new Exception("文档内容为空");
            _log($"文档读取完成，{text.Length} 个字符", false);

            // 2. AI 解析
            _log("正在 AI 解析订单内容...", false);
            var parser = new OrderParserService(ServiceLocator.Chat);
            var orderData = await parser.ParseOrderAsync(text);
            _log("AI 解析完成", false);

            // 保存 JSON 到同目录
            var jsonPath = Path.Combine(directory, Path.GetFileNameWithoutExtension(filePath) + ".json");
            await File.WriteAllTextAsync(jsonPath,
                JsonSerializer.Serialize(orderData.RootElement, new JsonSerializerOptions { WriteIndented = true }));
            _log($"JSON 已保存: {Path.GetFileName(jsonPath)}", false);

            // 3. 加载配置
            var config = await LoadConfigAsync();

            // 4. 浏览器自动化
            _log("正在启动浏览器...", false);
            await using var browser = new BrowserAutomationService(
                msg => _log(msg, false),
                _requestCaptcha);

            await browser.InitAsync();

            // 登录
            var systemUrl = config.GetValueOrDefault("SystemUrl", "http://192.168.1.12:8014");
            var sysUser = config.GetValueOrDefault("SystemUsername", "");
            var sysPass = config.GetValueOrDefault("SystemPassword", "");
            await browser.LoginAsync(systemUrl, sysUser, sysPass);

            // 导航到新增订单
            await browser.NavigateToNewOrderAsync();

            // 填写基础信息
            await browser.FillBasicInfoAsync(config, orderData);

            // 填写物品信息
            await browser.FillProductsAsync(orderData);

            // 上传原始文件
            await browser.UploadOrderFileAsync(filePath);

            _log($"✅ 订单填写完成: {fileName}", false);
            _log("请检查填写内容，确认无误后手动提交", false);

            // 等待用户确认后再移动文件（给30分钟）
            await Task.Delay(TimeSpan.FromMinutes(30));

            // 移动到已处理
            MoveFile(filePath, successDir, fileName);
            _log($"文件已移动到「已处理」", false);
        }
        catch (Exception ex)
        {
            _log($"❌ 处理失败 [{fileName}]: {ex.Message}", true);
            MoveFile(filePath, failDir, fileName);
            _log($"文件已移动到「失败」", true);
        }
        finally
        {
            lock (_lock) { _processingFiles.Remove(filePath); }
        }
    }

    private static void MoveFile(string source, string destDir, string fileName)
    {
        try
        {
            var dest = Path.Combine(destDir, fileName);
            if (File.Exists(dest))
                dest = Path.Combine(destDir, $"{Path.GetFileNameWithoutExtension(fileName)}_{DateTime.Now:yyyyMMddHHmmss}{Path.GetExtension(fileName)}");
            File.Move(source, dest);
        }
        catch { /* 忽略移动失败 */ }
    }

    private static async Task<Dictionary<string, string>> LoadConfigAsync()
    {
        await using var db = new AppDbContext();
        return await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);
    }

    private static bool IsWordFile(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".doc" or ".docx";
    }

    public void Dispose()
    {
        _watcher?.Dispose();
    }
}

using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using AIWorkAssistant.Services.Agent;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class OrderAgentViewModel : ObservableObject
{
    [ObservableProperty] private string _watchFolder = string.Empty;
    [ObservableProperty] private bool _isRunning;
    [ObservableProperty] private bool _isCaptchaVisible;
    [ObservableProperty] private byte[]? _captchaImageData;
    [ObservableProperty] private string _captchaInput = string.Empty;

    public ObservableCollection<AgentLogEntry> Logs { get; } = new();

    private OrderUploadAgent? _agent;
    private TaskCompletionSource<string>? _captchaTcs;

    public async Task LoadConfigAsync()
    {
        await using var db = new AppDbContext();
        var folder = (await db.AppSettings.FindAsync("WatchFolder"))?.Value ?? "";
        WatchFolder = folder;
    }

    [RelayCommand]
    private async Task StartAsync()
    {
        if (string.IsNullOrWhiteSpace(WatchFolder))
        {
            AddLog("请先设置监测文件夹路径", true);
            return;
        }

        // 确保浏览器内核已安装
        AddLog("检查浏览器内核...", false);
        try
        {
            await Task.Run(() => BrowserAutomationService.EnsureBrowserInstalled());
            AddLog("浏览器内核就绪", false);
        }
        catch (Exception ex)
        {
            AddLog($"浏览器内核安装失败: {ex.Message}", true);
            return;
        }

        _agent = new OrderUploadAgent(
            (msg, isError) => Avalonia.Threading.Dispatcher.UIThread.Post(() => AddLog(msg, isError)),
            RequestCaptchaAsync);

        _agent.Start(WatchFolder);
        IsRunning = true;

        // 保存文件夹路径
        await using var db = new AppDbContext();
        var existing = await db.AppSettings.FindAsync("WatchFolder");
        if (existing != null) existing.Value = WatchFolder;
        else db.AppSettings.Add(new AppSetting { Key = "WatchFolder", Value = WatchFolder });
        await db.SaveChangesAsync();
    }

    [RelayCommand]
    private void Stop()
    {
        _agent?.Stop();
        _agent?.Dispose();
        _agent = null;
        IsRunning = false;
    }

    [RelayCommand]
    private void SubmitCaptcha()
    {
        if (_captchaTcs != null && !string.IsNullOrWhiteSpace(CaptchaInput))
        {
            _captchaTcs.TrySetResult(CaptchaInput);
            CaptchaInput = "";
            IsCaptchaVisible = false;
        }
    }

    private Task<string> RequestCaptchaAsync(byte[] imageData)
    {
        _captchaTcs = new TaskCompletionSource<string>();

        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        {
            CaptchaImageData = imageData;
            IsCaptchaVisible = true;
            AddLog("⚠️ 请输入验证码", false);
        });

        return _captchaTcs.Task;
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

public class AgentLogEntry
{
    public DateTime Timestamp { get; set; }
    public string Message { get; set; } = string.Empty;
    public bool IsError { get; set; }
    public string Display => $"[{Timestamp:HH:mm:ss}] {Message}";
}

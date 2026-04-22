using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using AIWorkAssistant.Services;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class SettingsViewModel : ObservableObject
{
    // AI 配置
    [ObservableProperty] private string _apiBaseUrl = string.Empty;
    [ObservableProperty] private string _apiKey = string.Empty;
    [ObservableProperty] private string _model = string.Empty;
    [ObservableProperty] private string _statusMessage = string.Empty;

    // 用户管理
    public ObservableCollection<User> Users { get; } = new();

    // 助手管理
    public ObservableCollection<AiAssistant> Assistants { get; } = new();

    public bool IsAdmin => ServiceLocator.Auth.IsAdmin;

    public async Task LoadAsync()
    {
        await using var db = new AppDbContext();

        // 加载 AI 配置
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);
        ApiBaseUrl = settings.GetValueOrDefault("AiBaseUrl", "https://yxai.chat");
        ApiKey = settings.GetValueOrDefault("AiApiKey", "");
        Model = settings.GetValueOrDefault("AiModel", "glm-5.1");

        if (IsAdmin)
        {
            // 加载用户列表
            Users.Clear();
            foreach (var u in await db.Users.ToListAsync())
                Users.Add(u);

            // 加载助手列表
            Assistants.Clear();
            foreach (var a in await db.Assistants.ToListAsync())
                Assistants.Add(a);
        }
    }

    [RelayCommand]
    private async Task SaveAiConfigAsync()
    {
        await using var db = new AppDbContext();
        await UpsertSetting(db, "AiBaseUrl", ApiBaseUrl);
        await UpsertSetting(db, "AiApiKey", ApiKey);
        await UpsertSetting(db, "AiModel", Model);
        await db.SaveChangesAsync();

        // 更新运行时配置
        ServiceLocator.Chat.Configure(ApiBaseUrl, ApiKey, Model);
        StatusMessage = "✅ AI 配置已保存";
    }

    [RelayCommand]
    private async Task AddUserAsync()
    {
        // 占位：后续实现弹窗添加用户
        StatusMessage = "添加用户功能待实现";
    }

    [RelayCommand]
    private async Task AddAssistantAsync()
    {
        // 占位：后续实现弹窗添加助手
        StatusMessage = "添加助手功能待实现";
    }

    private static async Task UpsertSetting(AppDbContext db, string key, string value)
    {
        var existing = await db.AppSettings.FindAsync(key);
        if (existing != null)
        {
            existing.Value = value;
        }
        else
        {
            db.AppSettings.Add(new AppSetting { Key = key, Value = value });
        }
    }
}

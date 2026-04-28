using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class SettingsViewModel : ObservableObject
{
    private const string DefaultAiBaseUrl = "https://yxai.chat";
    private const string DefaultAiModel = "glm-5.1";

    // AI 配置
    [ObservableProperty] private string _apiBaseUrl = string.Empty;
    [ObservableProperty] private string _apiKey = string.Empty;
    [ObservableProperty] private string _model = string.Empty;

    [ObservableProperty] private string _statusMessage = string.Empty;

    public async Task LoadAsync()
    {
        await using var db = new AppDbContext();
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);

        ApiBaseUrl = FirstNonEmpty(settings.GetValueOrDefault("AiBaseUrl"), DefaultAiBaseUrl);
        ApiKey = settings.GetValueOrDefault("AiApiKey", "");
        Model = FirstNonEmpty(settings.GetValueOrDefault("AiModel"), DefaultAiModel);

    }

    [RelayCommand]
    private async Task SaveAllAsync()
    {
        await using var db = new AppDbContext();

        await UpsertSetting(db, "AiBaseUrl", ApiBaseUrl);
        await UpsertSetting(db, "AiApiKey", ApiKey);
        await UpsertSetting(db, "AiModel", Model);

        await db.SaveChangesAsync();

        Services.ServiceLocator.Chat.Configure(ApiBaseUrl, ApiKey, Model);
        StatusMessage = "✅ AI 配置已保存";
    }

    private static async Task UpsertSetting(AppDbContext db, string key, string value)
    {
        var existing = await db.AppSettings.FindAsync(key);
        if (existing != null)
            existing.Value = value ?? "";
        else
            db.AppSettings.Add(new AppSetting { Key = key, Value = value ?? "" });
    }

    private static string FirstNonEmpty(params string?[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
}

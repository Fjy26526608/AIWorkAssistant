using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class SettingsViewModel : ObservableObject
{
    // AI 配置
    [ObservableProperty] private string _apiBaseUrl = string.Empty;
    [ObservableProperty] private string _apiKey = string.Empty;
    [ObservableProperty] private string _model = string.Empty;

    // 订单上传系统配置
    [ObservableProperty] private string _systemUrl = string.Empty;
    [ObservableProperty] private string _systemUsername = string.Empty;
    [ObservableProperty] private string _systemPassword = string.Empty;
    [ObservableProperty] private string _defaultSaleType = string.Empty;
    [ObservableProperty] private string _defaultDeptName = string.Empty;
    [ObservableProperty] private string _defaultSaleManager = string.Empty;
    [ObservableProperty] private string _defaultMoneyType = string.Empty;
    [ObservableProperty] private string _defaultRate = string.Empty;
    [ObservableProperty] private string _defaultMarketType = string.Empty;

    [ObservableProperty] private string _statusMessage = string.Empty;

    public async Task LoadAsync()
    {
        await using var db = new AppDbContext();
        var settings = await db.AppSettings.ToDictionaryAsync(s => s.Key, s => s.Value);

        ApiBaseUrl = settings.GetValueOrDefault("AiBaseUrl", "https://yxai.chat");
        ApiKey = settings.GetValueOrDefault("AiApiKey", "");
        Model = settings.GetValueOrDefault("AiModel", "glm-5.1");

        SystemUrl = settings.GetValueOrDefault("SystemUrl", "http://192.168.1.12:8014");
        SystemUsername = settings.GetValueOrDefault("SystemUsername", "");
        SystemPassword = settings.GetValueOrDefault("SystemPassword", "");
        DefaultSaleType = settings.GetValueOrDefault("DefaultSaleType", "煤矿");
        DefaultDeptName = settings.GetValueOrDefault("DefaultDeptName", "矿用产品销售部");
        DefaultSaleManager = settings.GetValueOrDefault("DefaultSaleManager", "");
        DefaultMoneyType = settings.GetValueOrDefault("DefaultMoneyType", "人民币");
        DefaultRate = settings.GetValueOrDefault("DefaultRate", "1");
        DefaultMarketType = settings.GetValueOrDefault("DefaultMarketType", "老市场");
    }

    [RelayCommand]
    private async Task SaveAllAsync()
    {
        await using var db = new AppDbContext();

        await UpsertSetting(db, "AiBaseUrl", ApiBaseUrl);
        await UpsertSetting(db, "AiApiKey", ApiKey);
        await UpsertSetting(db, "AiModel", Model);

        await UpsertSetting(db, "SystemUrl", SystemUrl);
        await UpsertSetting(db, "SystemUsername", SystemUsername);
        await UpsertSetting(db, "SystemPassword", SystemPassword);
        await UpsertSetting(db, "DefaultSaleType", DefaultSaleType);
        await UpsertSetting(db, "DefaultDeptName", DefaultDeptName);
        await UpsertSetting(db, "DefaultSaleManager", DefaultSaleManager);
        await UpsertSetting(db, "DefaultMoneyType", DefaultMoneyType);
        await UpsertSetting(db, "DefaultRate", DefaultRate);
        await UpsertSetting(db, "DefaultMarketType", DefaultMarketType);

        await db.SaveChangesAsync();

        Services.ServiceLocator.Chat.Configure(ApiBaseUrl, ApiKey, Model);
        StatusMessage = "✅ 所有配置已保存";
    }

    private static async Task UpsertSetting(AppDbContext db, string key, string value)
    {
        var existing = await db.AppSettings.FindAsync(key);
        if (existing != null)
            existing.Value = value ?? "";
        else
            db.AppSettings.Add(new AppSetting { Key = key, Value = value ?? "" });
    }
}

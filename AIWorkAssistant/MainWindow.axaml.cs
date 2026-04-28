using Avalonia.Controls;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using AIWorkAssistant.ViewModels;
using AIWorkAssistant.Views.Pages;

namespace AIWorkAssistant;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        InitDatabase();
        // Login is temporarily bypassed; open the main workspace directly.
        ShowMain();
    }

    private void InitDatabase()
    {
        using var db = new AppDbContext();
        db.Database.EnsureCreated();
        SeedDefaultAssistants(db);
    }

    private static void SeedDefaultAssistants(AppDbContext db)
    {
        var changed = false;

        foreach (var legacyAssistant in db.Assistants.Where(a => a.Name == "订单自动上传"))
        {
            legacyAssistant.IsEnabled = false;
            changed = true;
        }

        if (!db.Assistants.Any(a => a.Name == "通用办公助手"))
        {
            db.Assistants.Add(new AiAssistant
            {
                Name = "通用办公助手",
                Description = "日常办公问答、文本润色和资料整理",
                Icon = "🤖",
                SystemPrompt = "你是一个专业、简洁的企业办公 AI 助手。",
                IsEnabled = true
            });
            changed = true;
        }

        if (changed)
        {
            db.SaveChanges();
        }
    }

    private void ShowLogin()
    {
        var vm = new LoginViewModel();
        vm.LoginSucceeded += ShowMain;
        PageHost.Content = new LoginPage { DataContext = vm };
    }

    private void ShowMain()
    {
        var vm = new MainViewModel();
        vm.LogoutRequested += ShowMain;

        vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainViewModel.CurrentPage))
            {
                if (vm.CurrentPage == "Settings")
                    ShowSettings(vm);
                else if (vm.CurrentPage == "OfficeOrderUpload")
                    ShowOfficeOrderUpload(vm);
            }
        };

        PageHost.Content = new MainPage { DataContext = vm };
        _ = vm.LoadAssistantsAsync();
    }

    private void ShowOfficeOrderUpload(MainViewModel mainVm)
    {
        var vm = new OfficeOrderUploadViewModel();
        var page = new OfficeOrderUploadPage { DataContext = vm };

        var container = new DockPanel();
        var backBtn = new Button { Content = "← 返回主页", Margin = new Avalonia.Thickness(12, 8) };
        backBtn.Click += (_, _) =>
        {
            mainVm.CurrentPage = "AssistantList";
            ShowMain();
        };
        DockPanel.SetDock(backBtn, Avalonia.Controls.Dock.Top);
        container.Children.Add(backBtn);
        container.Children.Add(page);

        PageHost.Content = container;
        _ = vm.LoadAsync();
    }

    private void ShowSettings(MainViewModel mainVm)
    {
        var vm = new SettingsViewModel();
        var page = new SettingsPage { DataContext = vm };

        var container = new DockPanel();
        var backBtn = new Button { Content = "← 返回", Margin = new Avalonia.Thickness(12, 8) };
        backBtn.Click += (_, _) =>
        {
            mainVm.CurrentPage = "AssistantList";
            ShowMain();
        };
        DockPanel.SetDock(backBtn, Avalonia.Controls.Dock.Top);
        container.Children.Add(backBtn);
        container.Children.Add(page);

        PageHost.Content = container;
        _ = vm.LoadAsync();
    }
}

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
        ShowLogin();
    }

    private void InitDatabase()
    {
        using var db = new AppDbContext();
        db.Database.EnsureCreated();
        SeedDefaultAssistants(db);
    }

    private static void SeedDefaultAssistants(AppDbContext db)
    {
        if (!db.Assistants.Any())
        {
            db.Assistants.Add(new AiAssistant
            {
                Name = "订单自动上传",
                Description = "监测文件夹，自动解析订单并上传到管理系统",
                Icon = "📋",
                SystemPrompt = "order_upload_agent",
                IsEnabled = true
            });
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
        vm.LogoutRequested += ShowLogin;

        vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainViewModel.CurrentPage))
            {
                if (vm.CurrentPage == "Settings")
                    ShowSettings(vm);
            }

            if (e.PropertyName == nameof(MainViewModel.SelectedAssistant))
            {
                if (vm.SelectedAssistant?.SystemPrompt == "order_upload_agent")
                {
                    ShowOrderAgent(vm);
                }
            }
        };

        PageHost.Content = new MainPage { DataContext = vm };
        _ = vm.LoadAssistantsAsync();
    }

    private void ShowOrderAgent(MainViewModel mainVm)
    {
        var vm = new OrderAgentViewModel();
        var page = new OrderAgentPage { DataContext = vm };

        var container = new DockPanel();
        var backBtn = new Button { Content = "← 返回助手列表", Margin = new Avalonia.Thickness(12, 8) };
        backBtn.Click += (_, _) =>
        {
            mainVm.BackToListCommand.Execute(null);
            ShowMain();
        };
        DockPanel.SetDock(backBtn, Avalonia.Controls.Dock.Top);
        container.Children.Add(backBtn);
        container.Children.Add(page);

        PageHost.Content = container;
        _ = vm.LoadConfigAsync();
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

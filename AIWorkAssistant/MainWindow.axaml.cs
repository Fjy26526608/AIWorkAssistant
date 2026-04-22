using Avalonia.Controls;
using AIWorkAssistant.Data;
using AIWorkAssistant.ViewModels;
using AIWorkAssistant.Views.Pages;
using Microsoft.EntityFrameworkCore;

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
    }

    private void ShowLogin()
    {
        var vm = new LoginViewModel();
        vm.LoginSucceeded += () =>
        {
            ShowMain();
        };

        var page = new LoginPage { DataContext = vm };
        PageHost.Content = page;
    }

    private void ShowMain()
    {
        var vm = new MainViewModel();
        vm.LogoutRequested += ShowLogin;

        // 监听页面切换
        vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainViewModel.CurrentPage))
            {
                if (vm.CurrentPage == "Settings")
                {
                    ShowSettings(vm);
                }
            }
        };

        var page = new MainPage { DataContext = vm };
        PageHost.Content = page;

        // 加载助手列表
        _ = vm.LoadAssistantsAsync();
    }

    private void ShowSettings(MainViewModel mainVm)
    {
        var vm = new SettingsViewModel();
        var page = new SettingsPage { DataContext = vm };

        // 用一个带返回按钮的容器
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

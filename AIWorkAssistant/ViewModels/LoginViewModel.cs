using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIWorkAssistant.Services;

namespace AIWorkAssistant.ViewModels;

public partial class LoginViewModel : ObservableObject
{
    [ObservableProperty] private string _username = string.Empty;
    [ObservableProperty] private string _password = string.Empty;
    [ObservableProperty] private string _errorMessage = string.Empty;
    [ObservableProperty] private bool _isLoading;

    public event Action? LoginSucceeded;

    [RelayCommand]
    private async Task LoginAsync()
    {
        ErrorMessage = "";

        if (string.IsNullOrWhiteSpace(Username) || string.IsNullOrWhiteSpace(Password))
        {
            ErrorMessage = "请输入用户名和密码";
            return;
        }

        IsLoading = true;
        try
        {
            var user = await ServiceLocator.Auth.LoginAsync(Username, Password);
            if (user != null)
            {
                LoginSucceeded?.Invoke();
            }
            else
            {
                ErrorMessage = "用户名或密码错误";
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = $"登录失败: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }
}

using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using AIWorkAssistant.Services;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.ViewModels;

public partial class MainViewModel : ObservableObject
{
    [ObservableProperty] private string _currentPage = "AssistantList";
    [ObservableProperty] private string _displayName = string.Empty;
    [ObservableProperty] private bool _isAdmin;
    [ObservableProperty] private AiAssistant? _selectedAssistant;

    public ObservableCollection<AiAssistant> Assistants { get; } = new();
    public ObservableCollection<ChatMessageVm> Messages { get; } = new();

    [ObservableProperty] private string _chatInput = string.Empty;
    [ObservableProperty] private bool _isSending;

    public event Action? LogoutRequested;

    public MainViewModel()
    {
        var user = ServiceLocator.Auth.CurrentUser;
        if (user != null)
        {
            DisplayName = user.DisplayName;
            IsAdmin = user.Role == "Admin";
        }
    }

    public async Task LoadAssistantsAsync()
    {
        Assistants.Clear();
        await using var db = new AppDbContext();
        var list = await db.Assistants.Where(a => a.IsEnabled).ToListAsync();
        foreach (var a in list)
            Assistants.Add(a);
    }

    [RelayCommand]
    private void SelectAssistant(AiAssistant assistant)
    {
        SelectedAssistant = assistant;
        Messages.Clear();
        CurrentPage = "Chat";
    }

    [RelayCommand]
    private async Task SendMessageAsync()
    {
        if (string.IsNullOrWhiteSpace(ChatInput) || SelectedAssistant == null) return;

        var userMsg = ChatInput.Trim();
        ChatInput = "";

        Messages.Add(new ChatMessageVm { Role = "user", Content = userMsg });
        IsSending = true;

        try
        {
            var history = Messages.Select(m => (m.Role, m.Content)).ToList();
            var reply = await ServiceLocator.Chat.SendMessageAsync(
                SelectedAssistant.SystemPrompt, history);

            Messages.Add(new ChatMessageVm { Role = "assistant", Content = reply });

            await using var db = new AppDbContext();
            var user = ServiceLocator.Auth.CurrentUser!;
            db.ChatMessages.Add(new ChatMessage
            {
                UserId = user.Id,
                AssistantId = SelectedAssistant.Id,
                Role = "user",
                Content = userMsg
            });
            db.ChatMessages.Add(new ChatMessage
            {
                UserId = user.Id,
                AssistantId = SelectedAssistant.Id,
                Role = "assistant",
                Content = reply
            });
            await db.SaveChangesAsync();
        }
        catch (Exception ex)
        {
            Messages.Add(new ChatMessageVm { Role = "assistant", Content = $"⚠️ 错误: {ex.Message}" });
        }
        finally
        {
            IsSending = false;
        }
    }

    [RelayCommand]
    private void BackToList()
    {
        SelectedAssistant = null;
        CurrentPage = "AssistantList";
    }

    [RelayCommand]
    private void GoToSettings()
    {
        CurrentPage = "Settings";
    }

    [RelayCommand]
    private void Logout()
    {
        ServiceLocator.Auth.Logout();
        LogoutRequested?.Invoke();
    }
}

public class ChatMessageVm
{
    public string Role { get; set; } = "user";
    public string Content { get; set; } = string.Empty;
    public bool IsUser => Role == "user";
}

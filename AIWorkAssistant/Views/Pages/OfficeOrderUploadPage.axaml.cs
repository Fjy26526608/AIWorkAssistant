using System.ComponentModel;
using AIWorkAssistant.ViewModels;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;
using Avalonia.Threading;

namespace AIWorkAssistant.Views.Pages;

public partial class OfficeOrderUploadPage : UserControl
{
    private OfficeOrderUploadViewModel? _subscribedViewModel;
    private ScrollViewer? _logScrollViewer;
    private TextBox? _logTextBox;

    public OfficeOrderUploadPage()
    {
        AvaloniaXamlLoader.Load(this);
        _logScrollViewer = this.FindControl<ScrollViewer>("LogScrollViewer");
        _logTextBox = this.FindControl<TextBox>("LogTextBox");
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object? sender, EventArgs e)
    {
        if (_subscribedViewModel != null)
        {
            _subscribedViewModel.PropertyChanged -= OnViewModelPropertyChanged;
        }

        _subscribedViewModel = DataContext as OfficeOrderUploadViewModel;
        if (_subscribedViewModel != null)
        {
            _subscribedViewModel.PropertyChanged += OnViewModelPropertyChanged;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(OfficeOrderUploadViewModel.LogText))
        {
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            if (_logTextBox == null || _logScrollViewer == null)
            {
                return;
            }

            _logTextBox.CaretIndex = _logTextBox.Text?.Length ?? 0;
            _logScrollViewer.ScrollToEnd();
        });
    }

    private async void BrowseFolderButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel?.StorageProvider == null)
        {
            return;
        }

        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "选择订单文件夹",
            AllowMultiple = false
        });

        if (folders.Count == 0 || DataContext is not OfficeOrderUploadViewModel vm)
        {
            return;
        }

        var localPath = folders[0].TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(localPath))
        {
            vm.OrderFolderPath = localPath;
        }
    }
}

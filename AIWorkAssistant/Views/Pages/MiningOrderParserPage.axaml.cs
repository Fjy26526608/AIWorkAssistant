using AIWorkAssistant.ViewModels;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;

namespace AIWorkAssistant.Views.Pages;

public partial class MiningOrderParserPage : UserControl
{
    public MiningOrderParserPage()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private async void BrowseFileButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel?.StorageProvider == null)
        {
            return;
        }

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "选择订单 Word/TXT 文件",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("订单文件")
                {
                    Patterns = new[] { "*.doc", "*.docx", "*.txt" }
                },
                FilePickerFileTypes.All
            }
        });

        if (files.Count == 0 || DataContext is not MiningOrderParserViewModel vm)
        {
            return;
        }

        var localPath = files[0].TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(localPath))
        {
            vm.FilePath = localPath;
        }
    }
}

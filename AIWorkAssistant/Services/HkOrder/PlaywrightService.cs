using Microsoft.Playwright;
using AIWorkAssistant.Models.HkOrder;

namespace AIWorkAssistant.Services.HkOrder;

public class PlaywrightService : IAsyncDisposable
{
    private const int DefaultViewportWidth = 1600;
    private const int DefaultViewportHeight = 1000;

    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private IPage? _page;

    public IPage? Page => _page;
    public bool IsRunning => _browser?.IsConnected == true;

    public async Task StartAsync(HkOrderSettings settings, Action<string>? log = null)
    {
        var browserName = settings.BrowserType.ToLower() switch
        {
            "firefox" => "firefox",
            "webkit"  => "webkit",
            _         => "chromium"
        };

        await Task.Run(() =>
        {
            log?.Invoke($"正在检查 {settings.BrowserType} 浏览器...");
            Microsoft.Playwright.Program.Main(["install", browserName]);
            log?.Invoke($"{settings.BrowserType} 浏览器就绪。");
        });

        _playwright = await Playwright.CreateAsync();

        var launchOptions = new BrowserTypeLaunchOptions
        {
            Headless = settings.Headless,
            SlowMo = settings.SlowMo
        };

        _browser = settings.BrowserType switch
        {
            "Firefox" => await _playwright.Firefox.LaunchAsync(launchOptions),
            "WebKit"  => await _playwright.Webkit.LaunchAsync(launchOptions),
            _         => await _playwright.Chromium.LaunchAsync(launchOptions)
        };

        var context = await _browser.NewContextAsync(new BrowserNewContextOptions
        {
            ViewportSize = new ViewportSize
            {
                Width = DefaultViewportWidth,
                Height = DefaultViewportHeight
            },
            ScreenSize = new ScreenSize
            {
                Width = DefaultViewportWidth,
                Height = DefaultViewportHeight
            }
        });
        context.SetDefaultTimeout(settings.DefaultTimeout);
        _page = await context.NewPageAsync();

        if (!string.IsNullOrWhiteSpace(settings.TargetUrl))
            await _page.GotoAsync(settings.TargetUrl, new() { WaitUntil = WaitUntilState.DOMContentLoaded });
    }

    public async Task StopAsync()
    {
        if (_browser != null)
        {
            await _browser.CloseAsync();
            _browser = null;
        }
        _playwright?.Dispose();
        _playwright = null;
        _page = null;
    }

    public async ValueTask DisposeAsync() => await StopAsync();
}

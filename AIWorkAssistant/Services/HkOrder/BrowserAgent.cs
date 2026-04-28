using System.IO;
using AIWorkAssistant.Models.HkOrder;
using Microsoft.Playwright;

namespace AIWorkAssistant.Services.HkOrder;

public class BrowserAgent(IPage page, HkOrderSettings settings, Action<string> log)
{
    private const int LoginRetryCount = 5;
    private const int LoginRedirectTimeoutMs = 8000;
    private const int ScreenshotViewportWidth = 1600;
    private const int ScreenshotViewportHeight = 1000;
    private const double ScreenshotZoom = 0.96;

    public async Task LoginAsync(Func<string, Task<string?>>? requestCaptchaInput = null)
    {
        log("[Agent] 步骤1：正在登录，等待验证码加载...");

        var captchaImg = page.Locator(
            "img[src*='captcha'], img[src*='code'], img[src*='verify'], .login-code img, .code-img img, img.captcha")
            .First;
        await captchaImg.WaitForAsync(new() { Timeout = 30000 });

        for (var attempt = 1; attempt <= LoginRetryCount; attempt++)
        {
            if (attempt > 1)
            {
                log($"[Agent] 登录未跳转，正在刷新验证码并重试 ({attempt}/{LoginRetryCount})...");
                await RefreshCaptchaAsync(captchaImg);
            }

            log($"[Agent] 验证码已加载，准备弹出输入窗口 ({attempt}/{LoginRetryCount})。");

            var captchaPath = Path.Combine(Path.GetTempPath(), $"captcha_{DateTime.Now:yyyyMMddHHmmss}_{attempt}.png");
            await captchaImg.ScreenshotAsync(new() { Path = captchaPath });

            var captchaText = requestCaptchaInput == null
                ? null
                : await requestCaptchaInput(captchaPath);

            if (string.IsNullOrWhiteSpace(captchaText))
            {
                throw new OperationCanceledException("验证码输入已取消。");
            }

            var passwordInput = page.Locator("input[type='password']").First;
            var formInputs = page.Locator("form input[type='text']:visible, form input:not([type]):visible, form input[type='tel']:visible");
            var usernameInput = formInputs.First;
            await usernameInput.FillAsync(settings.Username);
            await passwordInput.FillAsync(settings.Password);

            var captchaInput = await GetCaptchaInputAsync(captchaImg);
            await captchaInput.FillAsync(captchaText.Trim());
            await ClickLoginButtonAsync();
            log("[Agent] 已提交登录，请等待页面跳转...");

            if (await WaitForLoginSuccessAsync())
            {
                log("[Agent] 登录成功。");
                return;
            }
        }

        throw new InvalidOperationException("连续多次登录后页面仍未跳转，请检查账号、密码或验证码是否正确。");
    }

    public async Task NavigateToAddOrderAsync()
    {
        log("[Agent] 步骤2：跳转到新增订单页面...");

        await page.GotoAsync(settings.OrderPageUrl, new() { WaitUntil = WaitUntilState.NetworkIdle });
        await page.Locator(".el-form").First.WaitForAsync(new() { Timeout = 10000 });

        log("[Agent] 新增订单页面已加载。");
    }

    public async Task<string> FillFormAsync(OrderData order)
    {
        log("[Agent] 步骤3：填写订单表单...");

        var uploader = new OrderUploadService(page);
        var screenshotPath = await uploader.FillFormAsync(order, settings);

        log("[Agent] 表单填写完成。");
        return screenshotPath;
    }

    public async Task UploadDocFileAsync(string filePath)
    {
        log($"[Agent] 步骤4：上传订货通知单 {Path.GetFileName(filePath)}...");

        var fileInput = page.Locator("label:has-text('订货通知单')")
            .Locator("xpath=ancestor::div[contains(@class,'el-form-item')]")
            .Locator("input[type='file']");

        await fileInput.SetInputFilesAsync(filePath);
        await page.WaitForTimeoutAsync(500);

        log("[Agent] 文件已上传。");
    }

    public async Task SubmitAsync()
    {
        log("[Agent] 步骤5：提交订单...");

        var originalUrl = page.Url;
        await page.Locator("button.el-button--primary span:text('提交')").ClickAsync();
        var successMessage = await WaitForSubmitSuccessAsync(originalUrl);
        log($"[Agent] 提交成功：{successMessage}");
    }

    public async Task<string> TakeScreenshotAsync(string label = "")
    {
        await page.WaitForTimeoutAsync(1500);
        await PreparePageForScreenshotAsync();
        var screenshotBytes = await page.ScreenshotAsync(new PageScreenshotOptions { FullPage = true });
        var fileName = $"agent_{label}_{DateTime.Now:yyyyMMddHHmmss}.png";
        var path = Path.Combine(Path.GetTempPath(), fileName);
        await File.WriteAllBytesAsync(path, screenshotBytes);
        return path;
    }

    private async Task<ILocator> GetCaptchaInputAsync(ILocator captchaImg)
    {
        var directMatch = page.Locator(
            "input[placeholder*='验证码'], input[placeholder*='驗證碼'], input[name*='captcha'], input[name*='verify'], input[name*='code'], input[id*='captcha'], input[id*='verify']")
            .First;

        if (await directMatch.CountAsync() > 0)
        {
            return directMatch;
        }

        var form = captchaImg.Locator("xpath=ancestor::form[1]");
        var formInputs = form.Locator(
            "input:visible:not([type='password']):not([type='hidden']):not([type='submit']):not([readonly])");
        if (await formInputs.CountAsync() > 0)
        {
            return formInputs.Last;
        }

        var pageInputs = page.Locator(
            "input:visible:not([type='password']):not([type='hidden']):not([type='submit']):not([readonly])");
        if (await pageInputs.CountAsync() > 0)
        {
            return pageInputs.Last;
        }

        throw new InvalidOperationException("未找到验证码输入框。");
    }

    private async Task ClickLoginButtonAsync()
    {
        var loginButton = page.Locator(
            "button:has-text('登录'), button:has-text('登 录'), .login-btn, .el-button--primary, input[type='submit']")
            .First;

        if (await loginButton.CountAsync() == 0)
        {
            throw new InvalidOperationException("未找到登录按钮。");
        }

        await loginButton.ClickAsync();
    }

    private async Task<bool> WaitForLoginSuccessAsync()
    {
        try
        {
            await page.WaitForURLAsync(
                url => !url.Contains("login", StringComparison.OrdinalIgnoreCase),
                new() { Timeout = LoginRedirectTimeoutMs });
            return true;
        }
        catch (TimeoutException)
        {
            return false;
        }
        catch (PlaywrightException ex) when (ex.Message.Contains("Timeout", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }
    }

    private async Task RefreshCaptchaAsync(ILocator captchaImg)
    {
        await captchaImg.ClickAsync();
        await page.WaitForTimeoutAsync(800);
    }

    private async Task<string> WaitForSubmitSuccessAsync(string originalUrl)
    {
        var successLocator = page
            .Locator(".el-message:visible, .el-notification:visible, .el-message-box:visible, .el-dialog:visible, [role='alert']:visible")
            .Filter(new() { HasTextString = "成功" })
            .First;
        var submitButton = page.Locator("button.el-button--primary span:text('提交')").First;

        for (var i = 0; i < 50; i++)
        {
            if (HasUrlChanged(originalUrl, page.Url))
            {
                return $"页面已跳转：{page.Url}";
            }

            if (await successLocator.CountAsync() > 0 && await successLocator.IsVisibleAsync())
            {
                var text = (await successLocator.InnerTextAsync()).Trim();
                return string.IsNullOrWhiteSpace(text) ? "页面已出现成功提示" : text;
            }

            if (await submitButton.CountAsync() == 0 || !await submitButton.IsVisibleAsync())
            {
                return $"提交后页面状态已变化，当前地址：{page.Url}";
            }

            await page.WaitForTimeoutAsync(200);
        }

        throw new InvalidOperationException("点击提交后既没有看到成功提示，页面也没有跳转，当前订单不算成功。请检查页面是否真正保存成功。");
    }

    private static bool HasUrlChanged(string originalUrl, string currentUrl)
    {
        if (string.IsNullOrWhiteSpace(originalUrl) || string.IsNullOrWhiteSpace(currentUrl))
        {
            return !string.Equals(originalUrl, currentUrl, StringComparison.OrdinalIgnoreCase);
        }

        var original = originalUrl.Trim().TrimEnd('/');
        var current = currentUrl.Trim().TrimEnd('/');
        return !string.Equals(original, current, StringComparison.OrdinalIgnoreCase);
    }

    private async Task PreparePageForScreenshotAsync()
    {
        await page.SetViewportSizeAsync(ScreenshotViewportWidth, ScreenshotViewportHeight);
        await page.EvaluateAsync(
            """
            zoom => {
                document.documentElement.style.zoom = String(zoom);
                document.body.style.zoom = String(zoom);
                window.scrollTo(0, 0);
            }
            """,
            ScreenshotZoom);
        await page.WaitForTimeoutAsync(300);
    }
}

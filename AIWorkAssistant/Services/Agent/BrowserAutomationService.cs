using System.Text.Json;
using Microsoft.Playwright;

namespace AIWorkAssistant.Services.Agent;

/// <summary>
/// 浏览器自动化：登录系统、填写订单
/// </summary>
public class BrowserAutomationService : IAsyncDisposable
{
    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private IPage? _page;

    private readonly Action<string> _log;
    private readonly Func<byte[], Task<string>> _requestCaptcha;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="log">日志回调</param>
    /// <param name="requestCaptcha">验证码请求回调：传入图片字节，返回用户输入的验证码</param>
    public BrowserAutomationService(Action<string> log, Func<byte[], Task<string>> requestCaptcha)
    {
        _log = log;
        _requestCaptcha = requestCaptcha;
    }

    /// <summary>
    /// 确保 Playwright 浏览器内核已安装
    /// </summary>
    public static void EnsureBrowserInstalled()
    {
        var exitCode = Microsoft.Playwright.Program.Main(new[] { "install", "chromium" });
        if (exitCode != 0)
            throw new Exception($"Playwright 浏览器安装失败，退出码: {exitCode}");
    }

    public async Task InitAsync()
    {
        _playwright = await Playwright.CreateAsync();
        _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Headless = false, // 显示浏览器窗口，方便用户看到操作过程
            SlowMo = 300      // 稍微慢一点，让用户能看清
        });
        _page = await _browser.NewPageAsync();
    }

    /// <summary>
    /// 登录系统
    /// </summary>
    public async Task LoginAsync(string url, string username, string password)
    {
        _log("正在打开登录页面...");
        await _page!.GotoAsync(url, new PageGotoOptions { WaitUntil = WaitUntilState.NetworkIdle });

        // 填写用户名
        _log("填写用户名...");
        var usernameInput = _page.Locator("input[placeholder='账号']");
        await usernameInput.FillAsync(username);

        // 填写密码
        _log("填写密码...");
        var passwordInput = _page.Locator("input[type='password'][placeholder='密码']");
        await passwordInput.FillAsync(password);

        // 获取验证码图片
        _log("等待输入验证码...");
        var captchaImg = _page.Locator(".login-code img");
        var captchaBytes = await captchaImg.ScreenshotAsync();

        // 请求用户输入验证码
        var captchaCode = await _requestCaptcha(captchaBytes);

        if (string.IsNullOrWhiteSpace(captchaCode))
            throw new Exception("验证码未输入");

        // 填写验证码
        var captchaInput = _page.Locator("input[placeholder='验证码']");
        await captchaInput.FillAsync(captchaCode);

        // 点击登录
        _log("点击登录...");
        var loginBtn = _page.Locator("button:has-text('登 录')");
        await loginBtn.ClickAsync();

        // 等待页面跳转（登录成功后 URL 会变化）
        await _page.WaitForURLAsync(url => !url.Contains("login"), new PageWaitForURLOptions
        {
            Timeout = 30000
        });

        _log("登录成功！");
    }

    /// <summary>
    /// 导航到销售订单管理页面并点击新增
    /// </summary>
    public async Task NavigateToNewOrderAsync()
    {
        _log("正在导航到销售订单管理页面...");

        // 等待页面加载完成
        await _page!.WaitForLoadStateAsync(LoadState.NetworkIdle);

        // 尝试通过菜单导航（Element UI 侧边栏菜单）
        // 先找"销售订单管理"菜单项
        var menuItem = _page.Locator("text=销售订单管理").First;
        if (await menuItem.IsVisibleAsync())
        {
            await menuItem.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }

        // 点击新增按钮
        _log("点击新增订单...");
        await _page.WaitForTimeoutAsync(1000);
        var addBtn = _page.Locator("button:has-text('新增')").First;
        await addBtn.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await _page.WaitForTimeoutAsync(1000);

        _log("已打开新增订单页面");
    }

    /// <summary>
    /// 填写订单基础信息
    /// </summary>
    public async Task FillBasicInfoAsync(Dictionary<string, string> config, JsonDocument orderData)
    {
        _log("开始填写基础信息...");
        var root = orderData.RootElement;

        // 销售类型（下拉选择）
        await SelectDropdownAsync("saleType", config.GetValueOrDefault("DefaultSaleType", "煤矿"));

        // 销售部门（下拉选择）
        await SelectDropdownAsync("deptId", config.GetValueOrDefault("DefaultDeptName", "矿用产品销售部"));

        // 等待销售经理列表加载
        await _page!.WaitForTimeoutAsync(500);

        // 销售经理（下拉选择，依赖部门）
        var saleManager = config.GetValueOrDefault("DefaultSaleManager", "");
        if (!string.IsNullOrEmpty(saleManager))
        {
            await SelectDropdownAsync("saleUserId", saleManager);
        }

        // 客户名称（搜索下拉）
        var customerName = GetJsonString(root, "customerName") ?? "";
        if (!string.IsNullOrEmpty(customerName))
        {
            _log($"搜索客户: {customerName}");
            await SearchSelectAsync("customerId", customerName);
        }

        // 下单日期（默认今天）
        _log("设置下单日期...");
        var dateInput = _page.Locator("label[for='orderTime']").Locator("..").Locator("input");
        await dateInput.ClickAsync();
        // 日期选择器中点击"今天"
        var todayBtn = _page.Locator(".el-picker-panel__footer button:has-text('今天')");
        if (await todayBtn.IsVisibleAsync())
            await todayBtn.ClickAsync();
        else
            await _page.Keyboard.PressAsync("Escape");

        // 币种
        await SelectDropdownAsync("moneyType", config.GetValueOrDefault("DefaultMoneyType", "人民币"));

        // 汇率
        var rateInput = _page.Locator("label[for='rate']").Locator("..").Locator("input");
        await rateInput.FillAsync(config.GetValueOrDefault("DefaultRate", "1"));

        // 新老市场
        await SelectDropdownAsync("marketType", config.GetValueOrDefault("DefaultMarketType", "老市场"));

        // 备注：写入合同金额等信息
        var contractAmount = GetJsonString(root, "contractAmount");
        var remark = $"合同金额: {contractAmount}";
        var remarkInput = _page.Locator("label[for='remark']").Locator("..").Locator("textarea, input");
        if (await remarkInput.CountAsync() > 0)
        {
            await remarkInput.First.FillAsync(remark);
        }

        _log("基础信息填写完成");
    }

    /// <summary>
    /// 填写物品信息表格
    /// </summary>
    public async Task FillProductsAsync(JsonDocument orderData)
    {
        _log("开始填写物品信息...");
        var root = orderData.RootElement;

        // 获取产品列表
        var products = new List<(string name, string spec, double quantity, double unitPrice)>();

        // 从 netProducts 提取柔性网产品
        if (root.TryGetProperty("netProducts", out var netProducts))
        {
            foreach (var p in netProducts.EnumerateArray())
            {
                var spec = GetJsonString(p, "spec") ?? "";
                var width = GetJsonDouble(p, "width");
                var length = GetJsonDouble(p, "length");
                var amount = GetJsonDouble(p, "amount");
                var name = $"柔性网 {spec}";
                products.Add((name, spec, 1, amount));
            }
        }

        // 从 products 提取配件
        if (root.TryGetProperty("products", out var productList))
        {
            foreach (var p in productList.EnumerateArray())
            {
                var name = GetJsonString(p, "name") ?? "";
                var spec = GetJsonString(p, "spec") ?? "";
                var qty = GetJsonDouble(p, "quantity");
                var price = GetJsonDouble(p, "unitPrice");
                if (!string.IsNullOrEmpty(name) && qty > 0)
                {
                    products.Add((name, spec, qty, price));
                }
            }
        }

        _log($"共 {products.Count} 个物品需要填写");

        foreach (var (name, spec, quantity, unitPrice) in products)
        {
            _log($"添加物品: {name}");

            // 点击添加物品按钮
            var addBtn = _page!.Locator("button:has-text('添加物品'), button:has-text('添加')").Last;
            await addBtn.ClickAsync();
            await _page.WaitForTimeoutAsync(500);

            // 在最后一行的物品名称列搜索
            var lastRow = _page.Locator("table tbody tr").Last;

            // 找到物品名称/编码输入框，输入搜索
            var nameInput = lastRow.Locator("input").First;
            await nameInput.ClickAsync();
            await nameInput.FillAsync(name);
            await _page.WaitForTimeoutAsync(1000);

            // 从下拉列表中选择第一个匹配项
            var dropdownItem = _page.Locator(".el-select-dropdown__item, .el-autocomplete-suggestion__list li").First;
            if (await dropdownItem.IsVisibleAsync())
            {
                await dropdownItem.ClickAsync();
                await _page.WaitForTimeoutAsync(500);
            }

            // 填写数量
            if (quantity > 0)
            {
                var qtyInputs = lastRow.Locator("input");
                // 数量通常在第几列需要根据实际情况调整
                var qtyInput = qtyInputs.Nth(9); // 幅宽后面是数量
                await qtyInput.FillAsync(quantity.ToString());
            }

            // 填写单价
            if (unitPrice > 0)
            {
                var priceInputs = lastRow.Locator("input");
                var priceInput = priceInputs.Nth(11); // 原币含税单价
                await priceInput.FillAsync(unitPrice.ToString());
            }

            await _page.WaitForTimeoutAsync(300);
        }

        _log("物品信息填写完成");
    }

    /// <summary>
    /// 上传订货通知单文件
    /// </summary>
    public async Task UploadOrderFileAsync(string filePath)
    {
        _log("上传订货通知单...");
        var fileInput = _page!.Locator("input[type='file']").Last;
        await fileInput.SetInputFilesAsync(filePath);
        await _page.WaitForTimeoutAsync(1000);
        _log("文件上传完成");
    }

    #region Helper Methods

    /// <summary>
    /// Element UI 下拉框选择
    /// </summary>
    private async Task SelectDropdownAsync(string forAttr, string optionText)
    {
        // 点击下拉框触发展开
        var label = _page!.Locator($"label[for='{forAttr}']");
        var wrapper = label.Locator("..").Locator(".el-select .el-input__inner");
        await wrapper.ClickAsync();
        await _page.WaitForTimeoutAsync(300);

        // 选择选项
        var option = _page.Locator($".el-select-dropdown__item:has-text('{optionText}')").First;
        if (await option.IsVisibleAsync())
        {
            await option.ClickAsync();
        }
        else
        {
            _log($"⚠️ 未找到选项: {optionText}");
            await _page.Keyboard.PressAsync("Escape");
        }

        await _page.WaitForTimeoutAsync(200);
    }

    /// <summary>
    /// Element UI 搜索下拉框
    /// </summary>
    private async Task SearchSelectAsync(string forAttr, string searchText)
    {
        var label = _page!.Locator($"label[for='{forAttr}']");
        var input = label.Locator("..").Locator(".el-select input.el-input__inner");
        await input.ClickAsync();
        await input.FillAsync(searchText);
        await _page.WaitForTimeoutAsync(1500); // 等待搜索结果

        var option = _page.Locator(".el-select-dropdown__item").First;
        if (await option.IsVisibleAsync())
        {
            await option.ClickAsync();
        }
        else
        {
            _log($"⚠️ 搜索 '{searchText}' 无结果");
            await _page.Keyboard.PressAsync("Escape");
        }

        await _page.WaitForTimeoutAsync(200);
    }

    private static string? GetJsonString(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var val) && val.ValueKind == JsonValueKind.String)
            return val.GetString();
        if (el.TryGetProperty(prop, out val) && val.ValueKind == JsonValueKind.Number)
            return val.GetRawText();
        return null;
    }

    private static double GetJsonDouble(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var val))
        {
            if (val.ValueKind == JsonValueKind.Number)
                return val.GetDouble();
            if (val.ValueKind == JsonValueKind.String && double.TryParse(val.GetString(), out var d))
                return d;
        }
        return 0;
    }

    #endregion

    public async ValueTask DisposeAsync()
    {
        if (_browser != null)
            await _browser.CloseAsync();
        _playwright?.Dispose();
    }
}

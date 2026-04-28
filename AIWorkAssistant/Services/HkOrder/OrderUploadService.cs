using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using AIWorkAssistant.Models.HkOrder;
using Microsoft.Playwright;

namespace AIWorkAssistant.Services.HkOrder;

public class OrderUploadService(IPage page)
{
    private const int ExistenceCheckAttempts = 12;
    private const int ExistenceCheckDelayMs = 250;

    private static string EscapeTextSelector(string value) =>
        value.Replace("\\", "\\\\").Replace("\"", "\\\"");

    private static string NormalizeText(string? value) =>
        Regex.Replace(value ?? string.Empty, @"\s+", string.Empty);

    private ILocator GetFormItem(string labelText)
    {
        var escapedLabel = EscapeTextSelector(labelText);
        var label = page.Locator($".el-form-item__label:text-is(\"{escapedLabel}\")");
        return label.Locator(
            "xpath=ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' el-form-item ')][1]");
    }

    private ILocator GetTableRootFromRow(ILocator row) =>
        row.Locator(
            "xpath=ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' el-table ')][1]");

    private async Task<bool> WaitForLocatorAsync(ILocator locator, int attempts = ExistenceCheckAttempts, int delayMs = ExistenceCheckDelayMs)
    {
        for (var i = 0; i < attempts; i++)
        {
            if (await locator.CountAsync() > 0)
            {
                return true;
            }

            await page.WaitForTimeoutAsync(delayMs);
        }

        return false;
    }

    public async Task<string> FillFormAsync(OrderData order, HkOrderSettings settings)
    {
        if (string.IsNullOrWhiteSpace(order.CustomerName))
        {
            throw new InvalidOperationException("AI解析结果中客户名称为空，无法选择客户。请先检查 AI 解析输出。");
        }

        if (order.Items.Count == 0)
        {
            throw new InvalidOperationException("AI解析结果中没有物品，无法新增物品。请先检查 AI 解析输出。");
        }

        if (string.IsNullOrWhiteSpace(order.SaleDept))
        {
            throw new InvalidOperationException("AI解析结果中的销售部门为空，无法填写订单。请先检查 AI 解析输出。");
        }

        if (string.IsNullOrWhiteSpace(order.SaleManager))
        {
            throw new InvalidOperationException("AI解析结果中的销售经理为空，无法填写订单。请先检查 AI 解析输出。");
        }

        await SelectElOption("销售类型", settings.SaleType);
        await SelectElOption("销售部门", order.SaleDept);
        await SelectElOption("销售经理", order.SaleManager);
        await SelectElOption("产品线", "织布");
        await SearchAndSelectCustomer(order.CustomerName);
        await FillDate("下单日期", order.OrderDate);
        await SelectElOption("币种", "人民币");
        await FillInput("汇率", "1");
        await SelectElOption("新老市场", "老市场");
        foreach (var item in order.Items)
        {
            await page.Locator("button.el-button--text span:text(\"+ 增加\")").ClickAsync();
            await page.WaitForTimeoutAsync(300);

            var lastRow = page.Locator(".el-table__body tbody tr").Last;
            await SelectMaterialForRow(lastRow, item);
            await FillTableSelectCell(lastRow, "是否免费使用", "否");
            await FillTableCell(lastRow, "幅宽(米)", item.Width);
            await FillTableCell(lastRow, "数量", ResolveQuantity(item));
            if (IsSectionThreeMainItem(item))
            {
                await FillTableCell(lastRow, "原币价税合计", ResolveLineAmount(item));
            }
            else
            {
                await FillTableCell(lastRow, "原币含税单价", ResolveUnitPrice(item));
            }
            await FillTableCell(lastRow, "税率(%)", item.TaxRate);
            await FillTableCell(lastRow, "预发货日期", ResolveDeliveryDate(item));
            await page.WaitForTimeoutAsync(500);
        }

        await ClearFormFocusAsync();
        await page.WaitForTimeoutAsync(1500);
        var screenshotBytes = await page.ScreenshotAsync(new PageScreenshotOptions { FullPage = true });
        var screenshotPath = Path.Combine(Path.GetTempPath(), $"order_preview_{DateTime.Now:yyyyMMddHHmmss}.png");
        await File.WriteAllBytesAsync(screenshotPath, screenshotBytes);
        return screenshotPath;
    }

    private async Task SelectElOption(string labelText, string optionText)
    {
        if (string.IsNullOrWhiteSpace(optionText))
        {
            return;
        }

        var formItem = GetFormItem(labelText);
        if (!await WaitForLocatorAsync(formItem))
        {
            throw new InvalidOperationException(BuildSelectionNotFoundMessage(labelText, optionText));
        }

        var input = formItem.Locator(".el-select .el-input__inner").First;
        if (!await WaitForLocatorAsync(input))
        {
            throw new InvalidOperationException($"字段“{labelText}”不是可选择的下拉框。");
        }

        await input.ClickAsync();

        var searchInputs = page.Locator("input.el-select__input:visible");
        if (await searchInputs.CountAsync() > 0)
        {
            await searchInputs.Last.FillAsync(optionText);
            await page.WaitForTimeoutAsync(300);
        }

        var option = await FindDropdownOptionAsync(optionText);
        if (option == null)
        {
            throw new InvalidOperationException(BuildSelectionNotFoundMessage(labelText, optionText));
        }

        await option.ClickAsync();
    }

    private async Task SearchAndSelectCustomer(string customerName)
    {
        if (string.IsNullOrWhiteSpace(customerName))
        {
            return;
        }

        var formItem = GetFormItem("客户名称");
        if (!await WaitForLocatorAsync(formItem))
        {
            throw new InvalidOperationException(BuildSelectionNotFoundMessage("客户名称", customerName));
        }

        var input = formItem.Locator(".el-input__inner").First;
        if (!await WaitForLocatorAsync(input))
        {
            throw new InvalidOperationException("字段“客户名称”不是可选择的输入框。");
        }

        await input.ClickAsync();
        await input.FillAsync(customerName);
        await page.WaitForTimeoutAsync(800);

        var option = await FindDropdownOptionAsync(customerName);
        if (option == null)
        {
            throw new InvalidOperationException(BuildSelectionNotFoundMessage("客户名称", customerName));
        }

        await option.ClickAsync();
    }

    private async Task FillInput(string labelText, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var formItem = GetFormItem(labelText);
        var input = formItem.Locator("input.el-input__inner, textarea.el-textarea__inner").First;
        await WaitForLocatorAsync(input);
        await input.FillAsync(value);
    }

    private async Task FillDate(string labelText, string dateValue)
    {
        if (string.IsNullOrWhiteSpace(dateValue))
        {
            return;
        }

        var formItem = GetFormItem(labelText);
        var input = formItem.Locator(".el-date-editor input").First;
        await WaitForLocatorAsync(input);
        await input.FillAsync(dateValue);
        await input.PressAsync("Enter");
    }

    private async Task SelectMaterialForRow(ILocator row, OrderItem item)
    {
        if (string.IsNullOrWhiteSpace(item.Spec))
        {
            throw new InvalidOperationException("AI解析结果中的物料规格为空，无法按规格搜索物料。");
        }

        var table = GetTableRootFromRow(row);
        var itemNameColumnIndex = await GetTableColumnIndexAsync(table, "物品名称");
        var itemCell = row.Locator("td").Nth(itemNameColumnIndex);
        var selectButton = itemCell.Locator("button:has-text('选择')").First;

        if (!await WaitForLocatorAsync(selectButton))
        {
            throw new InvalidOperationException("物品名称列没有可用的“选择”按钮。");
        }

        await selectButton.ClickAsync();

        var dialog = page.Locator(".el-dialog:visible").Last;
        await dialog.WaitForAsync(new() { Timeout = 10000 });

        await FillDialogInput(dialog, "规格型号：", item.Spec);

        await dialog.Locator("button:has-text('查询')").First.ClickAsync();
        await WaitForMaterialSearchResultAsync(dialog);

        var resultTable = dialog.Locator(".right .el-table").First;
        var matchedRow = await FindMaterialRowBySpecOnlyWithRetryAsync(dialog, resultTable, item);
        if (matchedRow == null)
        {
            throw new InvalidOperationException(BuildMaterialNotFoundMessage(item));
        }

        await matchedRow.Locator("label.el-checkbox").First.ClickAsync();
        await page.WaitForTimeoutAsync(200);

        if (await dialog.IsVisibleAsync())
        {
            var primaryConfirmButton = dialog.Locator(".el-dialog__footer button.el-button--primary").First;
            var textConfirmButtons = dialog.Locator(".el-dialog__footer button:has-text('确定'), .el-dialog__footer button:has-text('确 定'), .el-dialog__footer button:has-text('确认')");

            if (await primaryConfirmButton.CountAsync() > 0)
            {
                await primaryConfirmButton.ClickAsync();
            }
            else if (await textConfirmButtons.CountAsync() > 0)
            {
                await textConfirmButtons.First.ClickAsync();
            }
            else
            {
                await matchedRow.DblClickAsync();
                await page.WaitForTimeoutAsync(300);

                if (await dialog.IsVisibleAsync())
                {
                    await page.Keyboard.PressAsync("Enter");
                }
            }
        }

        if (await dialog.IsVisibleAsync())
        {
            await dialog.WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = 10000 });
        }

        await page.WaitForTimeoutAsync(300);
    }

    private async Task FillDialogInput(ILocator dialog, string labelText, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var escapedLabel = EscapeTextSelector(labelText);
        var label = dialog.Locator($".el-form-item__label:text-is(\"{escapedLabel}\")").First;
        var formItem = label.Locator(
            "xpath=ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' el-form-item ')][1]");
        var input = formItem.Locator("input.el-input__inner, textarea.el-textarea__inner").First;
        await WaitForLocatorAsync(input);
        await input.FillAsync(value);
    }

    private async Task WaitForMaterialSearchResultAsync(ILocator dialog)
    {
        var resultTable = dialog.Locator(".right .el-table").First;
        var resultRows = resultTable.Locator(".el-table__body tbody tr");
        var emptyBlock = resultTable.Locator(".el-table__empty-block");
        var totalText = dialog.Locator(".el-pagination__total");

        await page.WaitForTimeoutAsync(700);

        for (var i = 0; i < 30; i++)
        {
            if (await resultRows.CountAsync() > 0)
            {
                return;
            }

            if (await totalText.CountAsync() > 0)
            {
                var text = await totalText.First.InnerTextAsync();
                if (Regex.IsMatch(text, @"共\s*[1-9]\d*\s*条"))
                {
                    return;
                }
            }

            if (i >= 8 &&
                await emptyBlock.CountAsync() > 0 &&
                await emptyBlock.First.IsVisibleAsync())
            {
                return;
            }

            await page.WaitForTimeoutAsync(250);
        }
    }

    private async Task<ILocator?> FindMaterialRowBySpecOnlyWithRetryAsync(ILocator dialog, ILocator table, OrderItem item)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            var matchedRow = await FindMaterialRowBySpecOnlyAsync(table, item);
            if (matchedRow != null)
            {
                return matchedRow;
            }

            await page.WaitForTimeoutAsync(400);

            if (attempt == 2)
            {
                await dialog.Locator("button:has-text('鏌ヨ')").First.ClickAsync();
                await WaitForMaterialSearchResultAsync(dialog);
            }
        }

        return null;
    }

    private async Task<int> GetTableColumnIndexAsync(ILocator table, string columnHeader)
    {
        var headers = table.Locator(".el-table__header th .cell");
        await WaitForLocatorAsync(headers);
        var count = await headers.CountAsync();
        var expected = NormalizeText(columnHeader);

        for (var i = 0; i < count; i++)
        {
            var text = NormalizeText(await headers.Nth(i).InnerTextAsync());
            if (text == expected)
            {
                return i;
            }
        }

        throw new InvalidOperationException($"未找到列：{columnHeader}");
    }

    private async Task<ILocator?> FindMaterialRowBySpecOnlyAsync(ILocator table, OrderItem item)
    {
        var specColumnIndex = await GetTableColumnIndexAsync(table, "规格型号");
        var rows = table.Locator(".el-table__body tbody tr");
        var count = await rows.CountAsync();
        if (count == 0)
        {
            return null;
        }

        var targetSpec = NormalizeText(item.Spec);
        ILocator? containsSpecRow = null;
        ILocator? firstRow = null;

        for (var i = 0; i < count; i++)
        {
            var currentRow = rows.Nth(i);
            firstRow ??= currentRow;

            if (string.IsNullOrWhiteSpace(targetSpec))
            {
                continue;
            }

            var specText = NormalizeText(await currentRow.Locator("td").Nth(specColumnIndex).InnerTextAsync());
            if (specText == targetSpec)
            {
                return currentRow;
            }

            if (containsSpecRow == null &&
                !string.IsNullOrWhiteSpace(specText) &&
                specText.Contains(targetSpec))
            {
                containsSpecRow = currentRow;
            }
        }

        return containsSpecRow ?? firstRow;
    }

    private async Task<ILocator?> FindMaterialRowAsync(ILocator table, OrderItem item)
    {
        var codeColumnIndex = await GetTableColumnIndexAsync(table, "物品编码");
        var nameColumnIndex = await GetTableColumnIndexAsync(table, "物品名称");
        var specColumnIndex = await GetTableColumnIndexAsync(table, "规格型号");

        var targetCode = NormalizeText(item.ItemCode);
        var targetName = NormalizeText(item.ItemName);
        var targetSpec = NormalizeText(item.Spec);

        ILocator? fallbackRow = null;

        var rows = table.Locator(".el-table__body tbody tr");
        var count = await rows.CountAsync();
        for (var i = 0; i < count; i++)
        {
            var currentRow = rows.Nth(i);
            var cells = currentRow.Locator("td");
            var codeText = NormalizeText(await cells.Nth(codeColumnIndex).InnerTextAsync());
            var nameText = NormalizeText(await cells.Nth(nameColumnIndex).InnerTextAsync());
            var specText = NormalizeText(await cells.Nth(specColumnIndex).InnerTextAsync());

            if (!string.IsNullOrWhiteSpace(targetCode) && codeText == targetCode)
            {
                return currentRow;
            }

            var nameMatched = !string.IsNullOrWhiteSpace(targetName) && nameText == targetName;
            var specMatched = string.IsNullOrWhiteSpace(targetSpec) || specText == targetSpec;
            if (nameMatched && specMatched)
            {
                return currentRow;
            }

            var partialNameMatched = !string.IsNullOrWhiteSpace(targetName) && nameText.Contains(targetName);
            var partialSpecMatched = string.IsNullOrWhiteSpace(targetSpec) || specText.Contains(targetSpec);
            if (fallbackRow == null && partialNameMatched && partialSpecMatched)
            {
                fallbackRow = currentRow;
            }
        }

        return fallbackRow;
    }

    private async Task FillTableCell(ILocator row, string colHeader, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var table = GetTableRootFromRow(row);
        var columnIndex = await GetTableColumnIndexAsync(table, colHeader);
        var cell = row.Locator("td").Nth(columnIndex);

        var dateInput = cell.Locator(".el-date-editor input").First;
        if (await WaitForLocatorAsync(dateInput, 3, 150))
        {
            await dateInput.ClickAsync();
            await dateInput.FillAsync(value);
            await dateInput.PressAsync("Enter");
            await page.Keyboard.PressAsync("Escape");
            await page.WaitForTimeoutAsync(150);
            return;
        }

        var input = cell.Locator("input, textarea").First;
        if (!await WaitForLocatorAsync(input))
        {
            throw new InvalidOperationException($"列“{colHeader}”不是可编辑输入框。");
        }

        await FillAndVerifyInputAsync(input, value, colHeader);
    }

    private async Task FillAndVerifyInputAsync(ILocator input, string value, string colHeader)
    {
        await input.ClickAsync();
        await input.PressAsync("Control+A");
        await input.FillAsync(string.Empty);
        await input.PressSequentiallyAsync(value, new() { Delay = 30 });
        await input.EvaluateAsync(
            """
            element => {
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
                if (typeof element.blur === "function") {
                    element.blur();
                }
            }
            """);
        await page.WaitForTimeoutAsync(120);

        var actualValue = await input.InputValueAsync();
        if (IsEquivalentInputValue(actualValue, value))
        {
            return;
        }

        await input.EvaluateAsync(
            """
            (element, targetValue) => {
                element.focus();
                element.value = targetValue;
                element.dispatchEvent(new Event('input', { bubbles: true }));
                element.dispatchEvent(new Event('change', { bubbles: true }));
                if (typeof element.blur === "function") {
                    element.blur();
                }
            }
            """,
            value);
        await page.WaitForTimeoutAsync(120);

        actualValue = await input.InputValueAsync();
        if (!IsEquivalentInputValue(actualValue, value))
        {
            throw new InvalidOperationException($"列“{colHeader}”填写失败，目标值“{value}”，实际值“{actualValue}”。");
        }
    }

    private static bool IsEquivalentInputValue(string? actualValue, string expectedValue)
    {
        var actual = actualValue?.Trim() ?? string.Empty;
        var expected = expectedValue.Trim();

        if (actual == expected)
        {
            return true;
        }

        var actualNumber = ParseDecimal(actual);
        var expectedNumber = ParseDecimal(expected);
        if (actualNumber.HasValue && expectedNumber.HasValue)
        {
            return actualNumber.Value == expectedNumber.Value;
        }

        return false;
    }

    private async Task FillTableSelectCell(ILocator row, string colHeader, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var table = GetTableRootFromRow(row);
        var columnIndex = await GetTableColumnIndexAsync(table, colHeader);
        var cell = row.Locator("td").Nth(columnIndex);
        var input = cell.Locator(".el-select .el-input__inner").First;
        if (!await WaitForLocatorAsync(input))
        {
            throw new InvalidOperationException($"列“{colHeader}”不是可选择的下拉框。");
        }

        await input.ClickAsync();

        var option = await FindDropdownOptionAsync(value);
        if (option == null)
        {
            throw new InvalidOperationException(BuildSelectionNotFoundMessage(colHeader, value));
        }

        await option.ClickAsync();
    }

    private async Task<ILocator?> FindDropdownOptionAsync(string optionText)
    {
        var escapedOption = EscapeTextSelector(optionText);
        var exactOption = page.Locator(
            $".el-select-dropdown__item:not(.is-disabled) span:text-is(\"{escapedOption}\")");
        var fuzzyOption = page.Locator(
            $".el-select-dropdown__item:not(.is-disabled) span:text(\"{escapedOption}\")");

        for (var i = 0; i < ExistenceCheckAttempts; i++)
        {
            if (await exactOption.CountAsync() > 0)
            {
                return exactOption.Last;
            }

            if (await fuzzyOption.CountAsync() > 0)
            {
                return fuzzyOption.Last;
            }

            await page.WaitForTimeoutAsync(ExistenceCheckDelayMs);
        }

        return null;
    }

    private static string BuildSelectionNotFoundMessage(string fieldName, string targetValue) =>
        $"选择“{fieldName}”时，未找到目标值“{targetValue}”，请检查后重试。";

    private static string BuildMaterialNotFoundMessage(OrderItem item)
    {
        var parts = new List<string>();

        if (!string.IsNullOrWhiteSpace(item.ItemCode))
        {
            parts.Add($"编码：{item.ItemCode}");
        }

        if (!string.IsNullOrWhiteSpace(item.ItemName))
        {
            parts.Add($"名称：{item.ItemName}");
        }

        if (!string.IsNullOrWhiteSpace(item.Spec))
        {
            parts.Add($"规格：{item.Spec}");
        }

        var materialInfo = parts.Count > 0 ? string.Join("，", parts) : "未提供物料信息";
        return $"物料搜索后不存在，请检查后重试。{materialInfo}";
    }

    private static string ResolveDeliveryDate(OrderItem item)
    {
        if (!string.IsNullOrWhiteSpace(item.DeliveryDate))
        {
            return item.DeliveryDate.Trim();
        }

        return DateTime.Today.AddDays(15).ToString("yyyy-MM-dd");
    }

    private static bool IsSectionThreeMainItem(OrderItem item)
    {
        if (string.IsNullOrWhiteSpace(item.Spec) ||
            string.IsNullOrWhiteSpace(item.Width) ||
            string.IsNullOrWhiteSpace(item.Length))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(item.LengthSegments))
        {
            return true;
        }

        var remark = item.ItemRemark?.Trim() ?? string.Empty;
        return remark.StartsWith("规格", StringComparison.Ordinal);
    }

    private static string ResolveQuantity(OrderItem item)
    {
        if (IsSectionThreeMainItem(item))
        {
            var width = ParseDecimal(item.Width);
            var length = ParseDecimal(item.Length);
            if (width.HasValue && length.HasValue && width.Value > 0 && length.Value > 0)
            {
                return (width.Value * length.Value).ToString("0.####", CultureInfo.InvariantCulture);
            }
        }

        return item.Quantity?.Trim() ?? string.Empty;
    }

    private static string ResolveUnitPrice(OrderItem item)
    {
        if (!string.IsNullOrWhiteSpace(item.UnitPrice))
        {
            return item.UnitPrice.Trim();
        }

        return string.IsNullOrWhiteSpace(item.Amount) ? string.Empty : item.Amount.Trim();
    }

    private static string ResolveLineAmount(OrderItem item)
    {
        if (!string.IsNullOrWhiteSpace(item.Amount))
        {
            return item.Amount.Trim();
        }

        return string.Empty;
    }

    private static decimal? ParseDecimal(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Trim()
            .Replace(",", string.Empty)
            .Replace("，", string.Empty)
            .Replace("￥", string.Empty)
            .Replace("¥", string.Empty);

        return decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out var invariantValue)
            ? invariantValue
            : decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out var currentValue)
                ? currentValue
                : null;
    }

    private async Task ClearFormFocusAsync()
    {
        await page.EvaluateAsync(
            """
            () => {
                if (document.activeElement && typeof document.activeElement.blur === "function") {
                    document.activeElement.blur();
                }
                window.scrollTo(0, 0);
            }
            """);
        var rateFormItem = GetFormItem("汇率");
        var rateInput = rateFormItem.Locator("input.el-input__inner").First;
        if (await rateInput.CountAsync() > 0)
        {
            await rateInput.ClickAsync();
        }
        await page.Keyboard.PressAsync("Escape");
        await page.Keyboard.PressAsync("Escape");
        await page.WaitForTimeoutAsync(200);
    }

    public async Task SubmitAsync()
    {
        await page.Locator("button.el-button--primary span:text(\"提交\")").ClickAsync();
    }
}

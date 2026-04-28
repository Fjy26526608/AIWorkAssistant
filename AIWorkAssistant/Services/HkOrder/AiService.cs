using System.Globalization;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using AIWorkAssistant.Models.HkOrder;

namespace AIWorkAssistant.Services.HkOrder;

public class AiService(string apiKey, string baseUrl, string model)
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(120) };

    private static readonly string[] ProductSectionStopKeywords =
    [
        "技术要求",
        "质量要求",
        "验收要求",
        "包装要求",
        "运输要求",
        "供货范围",
        "交货期",
        "交货日期",
        "付款方式",
        "结算方式",
        "其他要求",
        "备注说明"
    ];

    private static readonly string[] ExcludedProductLineKeywords =
    [
        "要求",
        "务必附",
        "布局图",
        "图纸中",
        "明确描述",
        "描述不能全面描述",
        "携带网卷",
        "携带钢丝绳"
    ];

    private const string SystemPrompt = """
        你是一个中文订单结构化提取助手。
        只返回 JSON，不要返回 markdown、解释、标题或任何额外文字。
        输出结构必须是：
        {
          "CustomerName": "",
          "SaleDept": "",
          "SaleManager": "",
          "OrderDate": "",
          "ContractAmount": "",
          "Remark": "",
          "Items": [
            {
              "ItemCode": "",
              "ItemName": "",
              "Spec": "",
              "Unit": "",
              "Width": "",
              "Length": "",
              "LengthSegments": "",
              "Quantity": "",
              "UnitPrice": "",
              "Amount": "",
              "TaxRate": "13",
              "DeliveryDate": "",
              "ItemRemark": ""
            }
          ]
        }

        规则：
        - CustomerName 优先提取“使用单位”。
        - SaleDept 优先提取“销售部门”等部门字段。
        - SaleManager 优先提取“销售经理”等人员字段。
        - ContractAmount 优先提取“合同金额”。
        - ItemName 必须保留订单原文中的叫法，不要擅自改名、归类、扩写。
        - 物品既包括“产品及配件”表格里的行，也包括正文里单独列出的物品行。
        - 第3部分里像“规格：... 金额：78000”这种规格行，就是主物品。
        - 对“产品 型号 数量 单价”表格，必须严格按列读取：第 3 列是 Quantity，第 4 列是 UnitPrice。
        - “整编钢丝绳：...”是物品；“整编钢丝绳要求：...”不是物品。
        - “携带网卷”“携带钢丝绳”等说明行不是物品。
        - 第3部分主物品按总价理解，金额写 Amount，不要把金额误当成单价乘数量。
        - 数量为 0 的物品不要输出。
        - 不知道就填空字符串。
        """;

    public async Task<AiParseResult> ParseOrderAsync(string docText, CancellationToken ct = default, Action<string>? log = null)
    {
        var itemSourceText = ExtractRelevantItemSections(docText);
        if (!string.IsNullOrWhiteSpace(itemSourceText))
        {
            var preview = itemSourceText[..Math.Min(300, itemSourceText.Length)];
            log?.Invoke($"  产品及配件片段预览: {preview}");
        }
        else
        {
            log?.Invoke("  未定位到“产品及配件”片段，将结合全文继续解析。");
        }

        var url = baseUrl.TrimEnd('/') + "/v1/messages";
        log?.Invoke($"  API request: {url}, model: {model}");

        var body = new
        {
            model,
            max_tokens = 3072,
            system = SystemPrompt,
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = BuildUserPrompt(docText, itemSourceText)
                }
            }
        };

        var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");

        var response = await Http.SendAsync(request, ct);
        var responseBytes = await response.Content.ReadAsByteArrayAsync(ct);
        var responseJson = Encoding.UTF8.GetString(responseBytes);

        log?.Invoke($"  API status: {(int)response.StatusCode}, response preview: {responseJson[..Math.Min(200, responseJson.Length)]}");

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"API returned {(int)response.StatusCode}: {responseJson}");
        }

        var rawText = ExtractModelText(responseJson);
        if (string.IsNullOrWhiteSpace(rawText))
        {
            throw new Exception($"AI 返回中未找到可解析文本。原始响应前500字：{responseJson[..Math.Min(500, responseJson.Length)]}");
        }

        log?.Invoke($"  AI raw output preview: {rawText[..Math.Min(300, rawText.Length)]}");

        var cleanedJson = ExtractJsonObject(rawText);
        var data = JsonSerializer.Deserialize<OrderData>(cleanedJson, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }) ?? new OrderData();

        data = PostProcessOrderData(data, docText, itemSourceText, log);

        return new AiParseResult
        {
            Data = data,
            RawResponseJson = responseJson,
            RawModelText = rawText,
            CleanedJson = cleanedJson
        };
    }

    private static string BuildUserPrompt(string docText, string productSection)
    {
        var builder = new StringBuilder();
        builder.AppendLine("以下是订单全文。");
        builder.AppendLine("请提取 CustomerName、SaleDept、SaleManager、OrderDate、ContractAmount、Remark 和 Items。");
        builder.AppendLine("Items 不能只看最下面的“产品及配件”列表，也要看全文中的真实物品行。");
        builder.AppendLine("重点：");
        builder.AppendLine("1. 第3部分里的“规格：H PET 700-700MS 宽15米 长度分1段 金额：78000”这种规格行就是主物品。");
        builder.AppendLine("2. “整编钢丝绳：Φ21.5 钢丝绳 30米 金额：500”是物品。");
        builder.AppendLine("3. “整编钢丝绳要求：...”和“携带网卷...”“携带钢丝绳...”不是物品。");
        builder.AppendLine("4. 对“产品 型号 数量 单价”表格，严格按列读取，数量和单价不要读错。");
        builder.AppendLine("5. ItemName 必须保持订单原文中的叫法，不要自己改名，不要固定写成“柔性网”。");
        builder.AppendLine("6. 第3部分主物品按总价理解，金额放 Amount；第13部分表格按数量和单价理解。");
        builder.AppendLine();
        builder.AppendLine("【订单全文】");
        builder.AppendLine(docText);

        if (!string.IsNullOrWhiteSpace(productSection))
        {
            builder.AppendLine();
            builder.AppendLine("【产品及配件片段】");
            builder.AppendLine(productSection);
        }

        return builder.ToString();
    }

    private static OrderData PostProcessOrderData(OrderData data, string docText, string productSection, Action<string>? log)
    {
        data.CustomerName = data.CustomerName?.Trim() ?? string.Empty;
        data.SaleDept = data.SaleDept?.Trim() ?? string.Empty;
        data.SaleManager = data.SaleManager?.Trim() ?? string.Empty;
        data.OrderDate = data.OrderDate?.Trim() ?? string.Empty;
        data.ContractAmount = data.ContractAmount?.Trim() ?? string.Empty;
        data.Remark = data.Remark?.Trim() ?? string.Empty;

        var validProductLines = ExtractValidItemLines(docText, productSection);
        var filteredItems = new List<OrderItem>();

        foreach (var rawItem in data.Items ?? [])
        {
            var item = NormalizeItem(rawItem);

            if (IsEmptyItem(item))
            {
                continue;
            }

            if (IsZeroQuantity(item.Quantity))
            {
                continue;
            }

            if (ShouldExcludeItem(item))
            {
                log?.Invoke($"  已忽略说明性行: 名称={item.ItemName}, 规格={item.Spec}, 数量={item.Quantity}");
                continue;
            }

            if (validProductLines.Count > 0 &&
                !IsLikelyProductTableItem(item) &&
                !IsItemGroundedInCandidateLines(item, validProductLines))
            {
                log?.Invoke($"  AI 行未在原文候选物品行中命中，已忽略: 名称={item.ItemName}, 规格={item.Spec}");
                continue;
            }

            filteredItems.Add(item);
        }

        var sourceItems = ExtractDeterministicItems(docText, productSection, log);
        var mergedItems = MergeWithDeterministicItems(filteredItems, sourceItems, log);
        ApplySectionThreeMainItemQuantities(mergedItems, log);
        ApplyDefaultQuantity(mergedItems, log);

        FillSingleMissingPriceFromContract(mergedItems, data.ContractAmount, log);
        data.Items = mergedItems;
        return data;
    }

    private static List<OrderItem> ExtractDeterministicItems(string docText, string productSection, Action<string>? log)
    {
        var items = new List<OrderItem>();
        var sourceText = productSection;
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            return items;
        }

        var lines = sourceText.Replace("\r\n", "\n").Split('\n');
        var currentItemName = string.Empty;
        var currentLength = string.Empty;
        var inProductTable = false;
        var tableRowBuffer = new List<string>();

        for (var i = 0; i < lines.Length; i++)
        {
            var line = NormalizeSourceLine(lines[i]);
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var contentLine = StripLeadingSerial(line);
            var nextLine = i + 1 < lines.Length
                ? StripLeadingSerial(NormalizeSourceLine(lines[i + 1]))
                : string.Empty;
            if (string.IsNullOrWhiteSpace(contentLine))
            {
                continue;
            }

            if (contentLine.Contains("柔性网尺寸", StringComparison.Ordinal))
            {
                currentLength = ExtractNumber(contentLine, @"长度\s*(?<value>\d+(?:\.\d+)?)\s*米");
                continue;
            }

            if (contentLine.Contains("产品及配件", StringComparison.Ordinal))
            {
                inProductTable = false;
                tableRowBuffer.Clear();
                continue;
            }

            if (IsProductTableHeader(contentLine) || IsSplitProductTableHeader(contentLine, nextLine))
            {
                inProductTable = true;
                tableRowBuffer.Clear();
                if (IsSplitProductTableHeader(contentLine, nextLine))
                {
                    i = SkipSplitProductTableHeader(lines, i);
                }
                continue;
            }

            if (inProductTable)
            {
                if (IsProductTableEnd(contentLine))
                {
                    FlushSplitProductTableBuffer(tableRowBuffer, items);
                    inProductTable = false;
                    tableRowBuffer.Clear();
                }
                else
                {
                    var tableItem = TryParseProductTableRow(contentLine);
                    if (tableItem != null)
                    {
                        FlushSplitProductTableBuffer(tableRowBuffer, items);
                        items.Add(tableItem);
                    }
                    else if (!IsSplitProductTableHeaderWord(contentLine))
                    {
                        tableRowBuffer.Add(contentLine);
                        if (tableRowBuffer.Count >= 4)
                        {
                            FlushSplitProductTableBuffer(tableRowBuffer, items);
                        }
                    }
                }

                continue;
            }

            var specItem = TryParseSpecItem(contentLine, currentItemName, currentLength);
            if (specItem != null)
            {
                items.Add(specItem);
                continue;
            }

            var ropeItem = TryParseSteelRopeItem(contentLine);
            if (ropeItem != null)
            {
                items.Add(ropeItem);
                continue;
            }

            var transportFrameItem = TryParseTransportFrameItem(contentLine, nextLine);
            if (transportFrameItem != null)
            {
                items.Add(transportFrameItem);
            }
        }

        if (items.Count > 0)
        {
            log?.Invoke($"  Source fallback parsed {items.Count} item(s).");
        }

        return items;
    }

    private static List<OrderItem> MergeWithDeterministicItems(List<OrderItem> aiItems, List<OrderItem> sourceItems, Action<string>? log)
    {
        if (sourceItems.Count == 0)
        {
            return aiItems;
        }

        var remainingAiItems = aiItems
            .Select(NormalizeItem)
            .ToList();
        var mergedItems = new List<OrderItem>();

        foreach (var sourceItem in sourceItems)
        {
            var matchIndex = FindBestItemMatchIndex(remainingAiItems, sourceItem);
            if (matchIndex >= 0)
            {
                var merged = MergeItems(remainingAiItems[matchIndex], sourceItem);
                mergedItems.Add(merged);
                remainingAiItems.RemoveAt(matchIndex);
            }
            else
            {
                mergedItems.Add(NormalizeItem(sourceItem));
            }
        }

        var keptUnmatchedAiItems = remainingAiItems
            .Where(IsLikelyProductTableItem)
            .Select(NormalizeItem)
            .ToList();

        if (keptUnmatchedAiItems.Count > 0)
        {
            mergedItems.AddRange(keptUnmatchedAiItems);
            log?.Invoke($"  Kept unmatched AI table items: {keptUnmatchedAiItems.Count}");
            foreach (var item in keptUnmatchedAiItems)
            {
                log?.Invoke($"    kept AI table item: 名称={item.ItemName}, 规格={item.Spec}, 数量={item.Quantity}, 单价={item.UnitPrice}");
            }
        }

        var droppedUnmatchedAiItems = remainingAiItems.Count - keptUnmatchedAiItems.Count;
        if (droppedUnmatchedAiItems > 0)
        {
            log?.Invoke($"  Dropped unmatched AI items: {droppedUnmatchedAiItems}");
            foreach (var item in remainingAiItems.Where(item => !IsLikelyProductTableItem(item)))
            {
                log?.Invoke($"    dropped AI item: 名称={item.ItemName}, 规格={item.Spec}, 数量={item.Quantity}, 单价={item.UnitPrice}, 金额={item.Amount}");
            }
        }

        var finalItems = mergedItems
            .Select(NormalizeItem)
            .Where(item => !IsEmptyItem(item))
            .Where(item => !IsZeroQuantity(item.Quantity))
            .ToList();

        log?.Invoke($"  Final merged items: {finalItems.Count}");
        return finalItems;
    }

    private static int FindBestItemMatchIndex(List<OrderItem> candidates, OrderItem sourceItem)
    {
        var sourceName = NormalizeForMatch(sourceItem.ItemName);
        var sourceSpec = NormalizeForMatch(sourceItem.Spec);
        var sourceCode = NormalizeForMatch(sourceItem.ItemCode);

        var bestIndex = -1;
        var bestScore = 0;

        for (var i = 0; i < candidates.Count; i++)
        {
            var candidate = candidates[i];
            var candidateName = NormalizeForMatch(candidate.ItemName);
            var candidateSpec = NormalizeForMatch(candidate.Spec);
            var candidateCode = NormalizeForMatch(candidate.ItemCode);

            var score = 0;

            if (!string.IsNullOrWhiteSpace(sourceCode) &&
                !string.IsNullOrWhiteSpace(candidateCode) &&
                sourceCode == candidateCode)
            {
                score += 1000;
            }

            if (!string.IsNullOrWhiteSpace(sourceSpec) &&
                !string.IsNullOrWhiteSpace(candidateSpec))
            {
                if (sourceSpec == candidateSpec)
                {
                    score += 400;
                }
                else if (sourceSpec.Contains(candidateSpec, StringComparison.Ordinal) ||
                         candidateSpec.Contains(sourceSpec, StringComparison.Ordinal))
                {
                    score += 220;
                }
            }

            if (!string.IsNullOrWhiteSpace(sourceName) &&
                !string.IsNullOrWhiteSpace(candidateName))
            {
                if (sourceName == candidateName)
                {
                    score += 160;
                }
                else if (sourceName.Contains(candidateName, StringComparison.Ordinal) ||
                         candidateName.Contains(sourceName, StringComparison.Ordinal))
                {
                    score += 100;
                }
            }

            if (score > bestScore)
            {
                bestScore = score;
                bestIndex = i;
            }
        }

        return bestScore >= 160 ? bestIndex : -1;
    }

    private static OrderItem MergeItems(OrderItem aiItem, OrderItem sourceItem)
    {
        var merged = NormalizeItem(aiItem);

        merged.ItemCode = PreferSourceValue(merged.ItemCode, sourceItem.ItemCode);
        merged.ItemName = PreferSourceValue(merged.ItemName, sourceItem.ItemName);
        merged.Spec = PreferSourceValue(merged.Spec, sourceItem.Spec);
        merged.Unit = PreferSourceValue(merged.Unit, sourceItem.Unit);
        merged.Width = PreferSourceValue(merged.Width, sourceItem.Width);
        merged.Length = PreferSourceValue(merged.Length, sourceItem.Length);
        merged.LengthSegments = PreferSourceValue(merged.LengthSegments, sourceItem.LengthSegments);
        merged.Quantity = PreferSourceValue(merged.Quantity, sourceItem.Quantity);
        merged.UnitPrice = PreferSourceValue(merged.UnitPrice, sourceItem.UnitPrice);
        merged.Amount = PreferSourceValue(merged.Amount, sourceItem.Amount);
        merged.TaxRate = PreferSourceValue(merged.TaxRate, sourceItem.TaxRate);
        merged.DeliveryDate = PreferSourceValue(merged.DeliveryDate, sourceItem.DeliveryDate);
        merged.ItemRemark = PreferSourceValue(merged.ItemRemark, sourceItem.ItemRemark);

        return merged;
    }

    private static string PreferSourceValue(string currentValue, string sourceValue) =>
        string.IsNullOrWhiteSpace(sourceValue) ? currentValue : sourceValue.Trim();

    private static string NormalizeSourceLine(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : value.Replace('\u3000', ' ').Trim();

    private static string StripLeadingSerial(string line)
    {
        var stripped = Regex.Replace(line, @"^\s*\d+\s+", string.Empty);
        stripped = Regex.Replace(stripped, @"^\s*[一二三四五六七八九十]+\s*[、.．]\s*", string.Empty);
        return stripped.Trim();
    }

    private static bool IsProductTableHeader(string line)
    {
        var normalized = NormalizeForMatch(line);
        return normalized.Contains("产品型号数量单价", StringComparison.Ordinal);
    }

    private static bool IsProductTableEnd(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return true;
        }

        if (line.Contains("其他要求", StringComparison.Ordinal) ||
            line.Contains("技术人员姓名及联系方式", StringComparison.Ordinal) ||
            line.Contains("为便于进一步沟通", StringComparison.Ordinal))
        {
            return true;
        }

        return Regex.IsMatch(line, @"^\d+\s+") && !Regex.IsMatch(line, @"^\d+\s*[xX＊*]");
    }

    private static bool IsSplitProductTableHeader(string line, string nextLine) =>
        NormalizeSourceLine(line) == "产品" &&
        NormalizeSourceLine(nextLine) == "型号";

    private static bool IsSplitProductTableHeaderWord(string line) =>
        line is "产品" or "型号" or "数量" or "单价";

    private static int SkipSplitProductTableHeader(string[] lines, int startIndex)
    {
        var index = startIndex;
        var expectedHeaders = new[] { "产品", "型号", "数量", "单价" };

        foreach (var header in expectedHeaders)
        {
            while (index < lines.Length)
            {
                var contentLine = StripLeadingSerial(NormalizeSourceLine(lines[index]));
                if (string.IsNullOrWhiteSpace(contentLine))
                {
                    index++;
                    continue;
                }

                if (contentLine == header)
                {
                    index++;
                    break;
                }

                return startIndex;
            }
        }

        return index - 1;
    }

    private static void FlushSplitProductTableBuffer(List<string> tableRowBuffer, List<OrderItem> items)
    {
        while (tableRowBuffer.Count >= 4)
        {
            var item = TryBuildTableItem(
                tableRowBuffer[0].Trim(),
                tableRowBuffer[1].Trim(),
                tableRowBuffer[2].Trim(),
                tableRowBuffer[3].Trim(),
                string.Join(" ", tableRowBuffer.Take(4)));

            if (item != null)
            {
                items.Add(item);
            }

            tableRowBuffer.RemoveRange(0, 4);
        }
    }

    private static OrderItem? TryParseSpecItem(string line, string currentItemName, string currentLength)
    {
        if (!Regex.IsMatch(line, @"^规格\s*[：:]"))
        {
            return null;
        }

        var spec = ExtractSpecValue(line);
        if (string.IsNullOrWhiteSpace(spec))
        {
            return null;
        }

        var width = ExtractNumber(line, @"宽\s*(?<value>\d+(?:\.\d+)?)\s*米");
        var lengthSegments = ExtractNumber(line, @"长度分\s*(?<value>\d+(?:\.\d+)?)\s*段");
        var amount = ExtractNumber(line, @"金额\s*[：:]\s*(?<value>\d+(?:\.\d+)?)");

        return NormalizeItem(new OrderItem
        {
            ItemName = currentItemName?.Trim() ?? string.Empty,
            Spec = spec,
            Width = width,
            Length = currentLength,
            LengthSegments = lengthSegments,
            UnitPrice = amount,
            Amount = amount,
            TaxRate = "13",
            ItemRemark = line
        });
    }

    private static string ExtractSpecValue(string line)
    {
        var match = Regex.Match(
            line,
            @"^规格\s*[：:]\s*(?<value>.+?)(?=\s+宽\s*\d|\s+长度分\s*\d|\s+金额\s*[：:]|$)");

        return match.Success ? match.Groups["value"].Value.Trim() : string.Empty;
    }

    private static OrderItem? TryParseSteelRopeItem(string line)
    {
        if (line.Contains("要求", StringComparison.Ordinal))
        {
            return null;
        }

        if (!Regex.IsMatch(line, @"^整编钢丝绳\s*[：:]"))
        {
            return null;
        }

        var payload = Regex.Replace(line, @"^整编钢丝绳\s*[：:]\s*", string.Empty);
        var amount = ExtractNumber(payload, @"金额\s*[：:]\s*(?<value>\d+(?:\.\d+)?)");
        var length = ExtractNumber(payload, @"(?<value>\d+(?:\.\d+)?)\s*米");
        var spec = Regex.Replace(payload, @"金额\s*[：:]\s*\d+(?:\.\d+)?", string.Empty);
        spec = Regex.Replace(spec, @"(?<value>\d+(?:\.\d+)?)\s*米", string.Empty).Trim(' ', '，', ',');

        return NormalizeItem(new OrderItem
        {
            ItemName = "整编钢丝绳",
            Spec = spec,
            Length = length,
            UnitPrice = amount,
            Amount = amount,
            TaxRate = "13",
            ItemRemark = line
        });
    }

    private static OrderItem? TryParseTransportFrameItem(string line, string nextLine)
    {
        if (!line.Contains("井下运输最大尺寸", StringComparison.Ordinal))
        {
            return null;
        }

        var length = ExtractNumber(line, @"长\s*(?<value>\d+(?:\.\d+)?)\s*m");
        var lengthValue = ParseDecimal(length);
        if (!lengthValue.HasValue)
        {
            return null;
        }

        var spec = ResolveTransportFrameSpec(lengthValue.Value);
        var amount = ExtractNumber(line, @"金额\s*[：:]\s*(?<value>\d+(?:\.\d+)?)");
        if (string.IsNullOrWhiteSpace(amount) &&
            !string.IsNullOrWhiteSpace(nextLine) &&
            nextLine.StartsWith("金额", StringComparison.Ordinal))
        {
            amount = ExtractNumber(nextLine, @"金额\s*[：:]\s*(?<value>\d+(?:\.\d+)?)");
        }

        return NormalizeItem(new OrderItem
        {
            ItemName = "运输框架",
            Spec = spec,
            Length = length,
            Amount = amount,
            TaxRate = "13",
            ItemRemark = line
        });
    }

    private static string ResolveTransportFrameSpec(decimal length)
    {
        if (length < 3)
        {
            return "长3米以下";
        }

        if (length < 5)
        {
            return "3m-5m（不含5m）";
        }

        if (length < 6)
        {
            return "5m-6m（不含6m）";
        }

        return "6m以上";
    }

    private static OrderItem? TryParseProductTableRow(string line)
    {
        var normalizedLine = NormalizeSourceLine(line);

        var columns = Regex
            .Split(normalizedLine.Trim(), @"\t+|\s{2,}")
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .ToArray();

        if (columns.Length >= 4)
        {
            var splitItem = TryBuildTableItem(
                string.Join(" ", columns.Take(columns.Length - 3)).Trim(),
                columns[^3].Trim(),
                columns[^2].Trim(),
                columns[^1].Trim(),
                normalizedLine);

            if (splitItem != null)
            {
                return splitItem;
            }
        }

        var tailMatch = Regex.Match(
            normalizedLine,
            @"^(?<prefix>.+?)\s+(?<quantity>\d+(?:\.\d+)?)\s+(?<unitPrice>\d+(?:\.\d+)?)$");

        if (!tailMatch.Success)
        {
            return null;
        }

        var prefix = tailMatch.Groups["prefix"].Value.Trim();
        var quantity = tailMatch.Groups["quantity"].Value.Trim();
        var unitPrice = tailMatch.Groups["unitPrice"].Value.Trim();
        if (ParseDecimal(quantity) == null || ParseDecimal(unitPrice) == null)
        {
            return null;
        }

        var prefixMatch = Regex.Match(prefix, @"^(?<name>.+?)\s+(?<spec>\S+)$");
        if (!prefixMatch.Success)
        {
            return null;
        }

        return TryBuildTableItem(
            prefixMatch.Groups["name"].Value.Trim(),
            prefixMatch.Groups["spec"].Value.Trim(),
            quantity,
            unitPrice,
            normalizedLine);
    }

    private static OrderItem? TryBuildTableItem(string name, string spec, string quantity, string unitPrice, string rawLine)
    {
        if (string.IsNullOrWhiteSpace(name) ||
            string.IsNullOrWhiteSpace(spec) ||
            ParseDecimal(quantity) == null ||
            ParseDecimal(unitPrice) == null)
        {
            return null;
        }

        return NormalizeItem(new OrderItem
        {
            ItemName = name,
            Spec = spec,
            Quantity = quantity,
            UnitPrice = unitPrice,
            TaxRate = "13",
            ItemRemark = rawLine
        });
    }

    private static string ExtractNumber(string value, string pattern)
    {
        var match = Regex.Match(value, pattern);
        return match.Success ? match.Groups["value"].Value.Trim() : string.Empty;
    }

    private static OrderItem NormalizeItem(OrderItem item) =>
        new()
        {
            ItemCode = item.ItemCode?.Trim() ?? string.Empty,
            ItemName = item.ItemName?.Trim() ?? string.Empty,
            Spec = item.Spec?.Trim() ?? string.Empty,
            Unit = item.Unit?.Trim() ?? string.Empty,
            Width = item.Width?.Trim() ?? string.Empty,
            Length = item.Length?.Trim() ?? string.Empty,
            LengthSegments = item.LengthSegments?.Trim() ?? string.Empty,
            Quantity = item.Quantity?.Trim() ?? string.Empty,
            UnitPrice = item.UnitPrice?.Trim() ?? string.Empty,
            Amount = item.Amount?.Trim() ?? string.Empty,
            TaxRate = string.IsNullOrWhiteSpace(item.TaxRate) ? "13" : item.TaxRate.Trim(),
            DeliveryDate = item.DeliveryDate?.Trim() ?? string.Empty,
            ItemRemark = item.ItemRemark?.Trim() ?? string.Empty
        };

    private static bool IsEmptyItem(OrderItem item) =>
        string.IsNullOrWhiteSpace(item.ItemCode) &&
        string.IsNullOrWhiteSpace(item.ItemName) &&
        string.IsNullOrWhiteSpace(item.Spec) &&
        string.IsNullOrWhiteSpace(item.Quantity) &&
        string.IsNullOrWhiteSpace(item.Amount) &&
        string.IsNullOrWhiteSpace(item.UnitPrice);

    private static bool IsZeroQuantity(string? quantity)
    {
        if (string.IsNullOrWhiteSpace(quantity))
        {
            return false;
        }

        var cleaned = quantity.Replace(",", string.Empty).Trim();
        return decimal.TryParse(cleaned, out var parsed) && parsed == 0;
    }

    private static void FillSingleMissingPriceFromContract(List<OrderItem> items, string contractAmount, Action<string>? log)
    {
        var contractTotal = ParseDecimal(contractAmount);
        if (contractTotal == null || items.Count == 0)
        {
            return;
        }

        var missingPriceItems = items
            .Where(item =>
                string.IsNullOrWhiteSpace(item.UnitPrice) &&
                string.IsNullOrWhiteSpace(item.Amount))
            .ToList();

        if (missingPriceItems.Count != 1)
        {
            return;
        }

        decimal knownTotal = 0;
        foreach (var item in items)
        {
            if (ReferenceEquals(item, missingPriceItems[0]))
            {
                continue;
            }

            var lineTotal = ResolveLineTotal(item);
            if (lineTotal.HasValue)
            {
                knownTotal += lineTotal.Value;
            }
        }

        var remaining = contractTotal.Value - knownTotal;
        if (remaining <= 0)
        {
            return;
        }

        var missingItem = missingPriceItems[0];
        var quantity = ParseDecimal(missingItem.Quantity);
        if (IsSectionThreeMainItem(missingItem))
        {
            missingItem.Amount = remaining.ToString("0.####", CultureInfo.InvariantCulture);
            log?.Invoke($"  已按合同总价反推第3部分主物品金额: 名称={missingItem.ItemName}, 金额={missingItem.Amount}");
            return;
        }

        if (quantity.HasValue && quantity.Value > 0)
        {
            missingItem.UnitPrice = (remaining / quantity.Value).ToString("0.####", CultureInfo.InvariantCulture);
        }
        else
        {
            missingItem.UnitPrice = remaining.ToString("0.####", CultureInfo.InvariantCulture);
        }

        log?.Invoke($"  已按合同总价反推缺失单价: 名称={missingItem.ItemName}, 单价={missingItem.UnitPrice}");
    }

    private static decimal? ResolveLineTotal(OrderItem item)
    {
        if (IsSectionThreeMainItem(item))
        {
            var mainItemAmount = ParseDecimal(item.Amount);
            if (mainItemAmount.HasValue)
            {
                return mainItemAmount.Value;
            }
        }

        var quantity = ParseDecimal(item.Quantity);
        var unitPrice = ParseDecimal(item.UnitPrice);
        if (unitPrice.HasValue)
        {
            return quantity.HasValue && quantity.Value > 0
                ? unitPrice.Value * quantity.Value
                : unitPrice.Value;
        }

        var amount = ParseDecimal(item.Amount);
        if (!amount.HasValue)
        {
            return null;
        }

        return amount.Value;
    }

    private static decimal? ParseDecimal(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Trim()
            .Replace(",", string.Empty)
            .Replace("￥", string.Empty)
            .Replace("¥", string.Empty);

        return decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out var invariantValue)
            ? invariantValue
            : decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out var currentValue)
                ? currentValue
                : null;
    }

    private static void ApplyDefaultQuantity(List<OrderItem> items, Action<string>? log)
    {
        foreach (var item in items)
        {
            if (!string.IsNullOrWhiteSpace(item.Quantity))
            {
                continue;
            }

            if (IsEmptyItem(item))
            {
                continue;
            }

            item.Quantity = "1";
            log?.Invoke($"  缺少数量，已默认补 1: 名称={item.ItemName}, 规格={item.Spec}");
        }
    }

    private static void ApplySectionThreeMainItemQuantities(List<OrderItem> items, Action<string>? log)
    {
        foreach (var item in items)
        {
            if (!IsSectionThreeMainItem(item))
            {
                continue;
            }

            var width = ParseDecimal(item.Width);
            var length = ParseDecimal(item.Length);
            if (!width.HasValue || !length.HasValue || width.Value <= 0 || length.Value <= 0)
            {
                continue;
            }

            var quantity = width.Value * length.Value;
            item.Quantity = quantity.ToString("0.####", CultureInfo.InvariantCulture);
            log?.Invoke($"  第3部分主物品数量已按长度*宽度计算: 名称={item.ItemName}, 规格={item.Spec}, 数量={item.Quantity}");
        }
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

    private static bool IsLikelyProductTableItem(OrderItem item)
    {
        if (ShouldExcludeItem(item))
        {
            return false;
        }

        return !string.IsNullOrWhiteSpace(item.ItemName) &&
               ParseDecimal(item.Quantity).HasValue &&
               ParseDecimal(item.UnitPrice).HasValue;
    }

    private static bool ShouldExcludeItem(OrderItem item)
    {
        var combined = $"{item.ItemName} {item.Spec} {item.ItemRemark}".Trim();
        if (string.IsNullOrWhiteSpace(combined))
        {
            return false;
        }

        foreach (var keyword in ExcludedProductLineKeywords)
        {
            if (combined.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static List<string> ExtractValidItemLines(string docText, string productSection)
    {
        var lines = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        void AddCandidate(string rawLine)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line) || ShouldExcludeLine(line))
            {
                return;
            }

            var normalized = NormalizeForMatch(line);
            if (!string.IsNullOrWhiteSpace(normalized) && seen.Add(normalized))
            {
                lines.Add(normalized);
            }
        }

        if (!string.IsNullOrWhiteSpace(productSection))
        {
            foreach (var rawLine in productSection.Replace("\r\n", "\n").Split('\n'))
            {
                if (rawLine.Contains("产品及配件", StringComparison.Ordinal))
                {
                    continue;
                }

                AddCandidate(rawLine);
            }
        }

        return lines;
    }

    private static bool ShouldExcludeLine(string line)
    {
        foreach (var keyword in ExcludedProductLineKeywords)
        {
            if (line.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsLikelyStandaloneItemLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        if (ShouldExcludeLine(line))
        {
            return false;
        }

        if (line.Contains("合同金额", StringComparison.Ordinal) ||
            line.Contains("使用单位", StringComparison.Ordinal) ||
            line.Contains("其他要求", StringComparison.Ordinal))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"^规格\s*[：:]"))
        {
            return true;
        }

        if ((Regex.IsMatch(line, @"^整编钢丝绳\s*[：:]") && !line.Contains("要求", StringComparison.Ordinal)) ||
            line.Contains("井下运输最大尺寸", StringComparison.Ordinal))
        {
            return true;
        }

        return false;
    }

    private static bool IsItemGroundedInCandidateLines(OrderItem item, List<string> validProductLines)
    {
        var name = NormalizeForMatch(item.ItemName);
        var spec = NormalizeForMatch(item.Spec);
        var code = NormalizeForMatch(item.ItemCode);

        foreach (var line in validProductLines)
        {
            if (!string.IsNullOrWhiteSpace(code) && code.Length >= 4 && line.Contains(code, StringComparison.Ordinal))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(name) && name.Length >= 2 && line.Contains(name, StringComparison.Ordinal))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(spec) && spec.Length >= 2 && line.Contains(spec, StringComparison.Ordinal))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(name) &&
                !string.IsNullOrWhiteSpace(spec) &&
                line.Contains(name, StringComparison.Ordinal) &&
                line.Contains(spec, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeForMatch(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        return Regex.Replace(text, @"[\s\r\n\t:：;；、【】\[\]\(\)（）<>《》""'`]+", string.Empty);
    }

    private static string ExtractRelevantItemSections(string docText)
    {
        if (string.IsNullOrWhiteSpace(docText))
        {
            return string.Empty;
        }

        var lines = docText.Replace("\r\n", "\n").Split('\n');
        var targets = new HashSet<string>(StringComparer.Ordinal) { "3", "6", "13" };
        var sections = new List<string>();

        for (var i = 0; i < lines.Length; i++)
        {
            var rawLine = NormalizeSourceLine(lines[i]);
            if (string.IsNullOrWhiteSpace(rawLine))
            {
                continue;
            }

            var serial = ExtractLeadingSectionNumber(rawLine);
            if (serial == null || !targets.Contains(serial))
            {
                continue;
            }

            var buffer = new List<string> { rawLine };
            for (var j = i + 1; j < lines.Length; j++)
            {
                var nextLine = NormalizeSourceLine(lines[j]);
                if (string.IsNullOrWhiteSpace(nextLine))
                {
                    buffer.Add(string.Empty);
                    continue;
                }

                var nextSerial = ExtractLeadingSectionNumber(nextLine);
                if (nextSerial != null)
                {
                    break;
                }

                buffer.Add(nextLine);
            }

            sections.Add(string.Join('\n', buffer).Trim());
        }

        return string.Join("\n\n", sections.Where(section => !string.IsNullOrWhiteSpace(section)));
    }

    private static string? ExtractLeadingSectionNumber(string line)
    {
        var match = Regex.Match(line, @"^\s*(?<num>\d+)(?:\s+.*)?$");
        if (!match.Success)
        {
            return null;
        }

        if (!int.TryParse(match.Groups["num"].Value, out var number))
        {
            return null;
        }

        return number is >= 1 and <= 20
            ? match.Groups["num"].Value
            : null;
    }

    private static string ExtractProductAndPartsSection(string docText)
    {
        if (string.IsNullOrWhiteSpace(docText))
        {
            return string.Empty;
        }

        var normalizedText = docText.Replace("\r\n", "\n");
        var lines = normalizedText.Split('\n');
        var startIndex = Array.FindIndex(lines, line => line.Contains("产品及配件", StringComparison.Ordinal));
        if (startIndex < 0)
        {
            return string.Empty;
        }

        var collected = new List<string>();
        for (var i = startIndex; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (string.IsNullOrWhiteSpace(line))
            {
                if (collected.Count > 0)
                {
                    collected.Add(string.Empty);
                }

                continue;
            }

            if (i > startIndex && IsLikelyProductSectionBoundary(line))
            {
                break;
            }

            collected.Add(line);
            if (collected.Count >= 120)
            {
                break;
            }
        }

        return string.Join('\n', collected).Trim();
    }

    private static bool IsLikelyProductSectionBoundary(string line)
    {
        foreach (var keyword in ProductSectionStopKeywords)
        {
            if (line.Contains(keyword, StringComparison.Ordinal))
            {
                return true;
            }
        }

        if (Regex.IsMatch(line, @"^[一二三四五六七八九十\d]+[、.．]"))
        {
            return true;
        }

        if (Regex.IsMatch(line, @"^(备注|说明|技术|质量|包装|运输|验收|付款|结算|供货|交货|附图|图纸)\s*[:：]"))
        {
            return true;
        }

        return false;
    }

    private static string ExtractModelText(string responseJson)
    {
        using var doc = JsonDocument.Parse(responseJson);
        var root = doc.RootElement;

        if (root.TryGetProperty("content", out var anthropicContent) &&
            anthropicContent.ValueKind == JsonValueKind.Array)
        {
            var parts = new List<string>();
            foreach (var item in anthropicContent.EnumerateArray())
            {
                if (item.TryGetProperty("text", out var textNode) && textNode.ValueKind == JsonValueKind.String)
                {
                    parts.Add(textNode.GetString() ?? string.Empty);
                }
            }

            if (parts.Count > 0)
            {
                return string.Join("\n", parts).Trim();
            }
        }

        if (root.TryGetProperty("choices", out var choices) &&
            choices.ValueKind == JsonValueKind.Array &&
            choices.GetArrayLength() > 0)
        {
            var firstChoice = choices[0];

            if (firstChoice.TryGetProperty("message", out var message) &&
                message.TryGetProperty("content", out var contentNode))
            {
                var extracted = ExtractContentNodeText(contentNode);
                if (!string.IsNullOrWhiteSpace(extracted))
                {
                    return extracted.Trim();
                }
            }

            if (firstChoice.TryGetProperty("text", out var textNode) && textNode.ValueKind == JsonValueKind.String)
            {
                return (textNode.GetString() ?? string.Empty).Trim();
            }
        }

        if (root.TryGetProperty("output_text", out var outputText) && outputText.ValueKind == JsonValueKind.String)
        {
            return (outputText.GetString() ?? string.Empty).Trim();
        }

        if (root.TryGetProperty("response", out var responseNode) && responseNode.ValueKind == JsonValueKind.String)
        {
            return (responseNode.GetString() ?? string.Empty).Trim();
        }

        return string.Empty;
    }

    private static string ExtractContentNodeText(JsonElement contentNode)
    {
        if (contentNode.ValueKind == JsonValueKind.String)
        {
            return contentNode.GetString() ?? string.Empty;
        }

        if (contentNode.ValueKind == JsonValueKind.Array)
        {
            var parts = new List<string>();
            foreach (var item in contentNode.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                {
                    parts.Add(item.GetString() ?? string.Empty);
                    continue;
                }

                if (item.ValueKind == JsonValueKind.Object &&
                    item.TryGetProperty("text", out var textNode) &&
                    textNode.ValueKind == JsonValueKind.String)
                {
                    parts.Add(textNode.GetString() ?? string.Empty);
                }
            }

            return string.Join("\n", parts);
        }

        return string.Empty;
    }

    private static string ExtractJsonObject(string rawText)
    {
        var text = rawText.Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            text = text.Split('\n', 2)[1];
            text = text[..text.LastIndexOf("```", StringComparison.Ordinal)].Trim();
        }

        var start = text.IndexOf('{');
        var end = text.LastIndexOf('}');
        if (start < 0 || end < 0 || end <= start)
        {
            throw new Exception($"AI did not return valid JSON: {text[..Math.Min(300, text.Length)]}");
        }

        return text[start..(end + 1)];
    }
}

public class AiParseResult
{
    public OrderData Data { get; set; } = new();
    public string RawResponseJson { get; set; } = "";
    public string RawModelText { get; set; } = "";
    public string CleanedJson { get; set; } = "";
}

file class AnthropicResponse
{
    [JsonPropertyName("content")]
    public List<ContentBlock>? Content { get; set; }
}

file class ContentBlock
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }
}

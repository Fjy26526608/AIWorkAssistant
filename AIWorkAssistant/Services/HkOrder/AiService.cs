using System.Globalization;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using AIWorkAssistant.Models.HkOrder;

namespace AIWorkAssistant.Services.HkOrder;

public class AiService(string apiKey, string baseUrl, string model)
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(10) };
    private const int MainParseMaxTokens = 12000;
    private const int CatalogMatchMaxTokens = 4096;
    private const int CatalogMatchSourceMaxChars = 4000;
    private const int CatalogMatchCandidatesPerItem = 35;

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
          "FlexibleNet": "",
          "Spec": "",
          "BraidedSteelWireRope": "",
          "UndergroundTransportMaxLength": "",
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
        - 举例：“柔性网尺寸”中的“柔性网”是物品名称，不是尺寸字段。物品名称可能是多种多样的
        - 第3部分每一行“规格：...”都必须输出为一条 ItemName="xxx" 的明细。
        - 合同金额下面如果出现“规格：xxx 单价 n；规格：yyy 单价 m”，这些是第3部分主物品不同规格的单价。
        - 第3部分同一行里出现“规格：xxx 宽 a 米 然后规格：yyy 宽 b 米”时，必须拆成多条主物品明细。
        - 第3部分“混编：...”以及“携带网卷”“携带钢丝绳”是要求说明，不是物品明细。
        - FlexibleNet 提取“柔性网尺寸”整句或柔性网相关描述，例如“柔性网尺寸：长度256米”。
        - Spec 汇总“规格：...”行里的规格型号；多个规格用中文分号分隔。
        - BraidedSteelWireRope 提取“整编钢丝绳：...”行，不要提取“整编钢丝绳要求”。
        - UndergroundTransportMaxLength 只提取“井下运输最大尺寸”里的“长”，例如“6m”。
        - ItemName 必须保留订单原文中的叫法，不要擅自改名、归类、扩写。
        - 物品既包括“产品及配件”表格里的行，也包括正文里单独列出的物品行。
        - 第3部分里像“规格：... 金额：78000”这种规格行，就是主物品。
        - 对“产品 型号 数量 单价”表格，必须严格按列读取：第 3 列是 Quantity，第 4 列是 UnitPrice。
        - “整编钢丝绳：...”是物品；“整编钢丝绳要求：...”不是物品。
        - “携带网卷”“携带钢丝绳”等说明行不是物品。
        - 第3部分主物品按总价理解，金额写 Amount，不要把金额误当成单价乘数量。
        - 第3部分主物品若有宽度、长度和单价，则 Amount=宽度*长度*单价；没有单价且只有一个主物品缺少金额时，Amount 留空由程序用合同总金额反推。
        - 原文明写“金额”的物品，Amount 必须使用原文金额；只有 Quantity 和 UnitPrice 的物品，Amount 按 Quantity * UnitPrice 计算；整张订单只会有一个物品缺少金额。
        - 数量为 0 的物品不要输出。
        - 不知道就填空字符串。
        """;

    public async Task<AiParseResult> ParseOrderAsync(string docText, CancellationToken ct = default, Action<string>? log = null)
    {
        var itemSourceText = ExtractRelevantItemSections(docText);
        var structuredDraft = BuildStructuredDraft(docText);
        if (!string.IsNullOrWhiteSpace(itemSourceText))
        {
            var preview = itemSourceText[..Math.Min(300, itemSourceText.Length)];
            log?.Invoke($"  产品及配件片段预览: {preview}");
        }
        else
        {
            log?.Invoke("  未定位到“产品及配件”片段，将结合全文继续解析。");
        }

        if (!string.IsNullOrWhiteSpace(structuredDraft))
        {
            var preview = structuredDraft[..Math.Min(500, structuredDraft.Length)];
            log?.Invoke($"  预解析草稿预览: {preview}");
        }

        var url = baseUrl.TrimEnd('/') + "/v1/messages";
        log?.Invoke($"  API request: {url}, model: {model}");
        var userPrompt = BuildUserPrompt(docText, itemSourceText, structuredDraft);
        log?.Invoke($"  AI 解析请求大小: 用户prompt约 {userPrompt.Length:N0} 字，系统prompt约 {SystemPrompt.Length:N0} 字，max_tokens={MainParseMaxTokens}");

        var body = new
        {
            model,
            max_tokens = MainParseMaxTokens,
            stream = false,
            system = SystemPrompt,
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = userPrompt
                }
            }
        };

        var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

        HttpResponseMessage response;
        var stopwatch = Stopwatch.StartNew();
        try
        {
            response = await Http.SendAsync(request, ct);
        }
        catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
        {
            throw new TimeoutException("AI 解析请求超过 10 分钟仍未返回。请稍后重试，或检查网络、模型响应速度和订单文本长度。", ex);
        }

        var responseBytes = await response.Content.ReadAsByteArrayAsync(ct);
        stopwatch.Stop();
        var responseJson = Encoding.UTF8.GetString(responseBytes);

        log?.Invoke($"  API status: {(int)response.StatusCode}, 耗时 {stopwatch.Elapsed.TotalSeconds:F1} 秒, response preview: {responseJson[..Math.Min(200, responseJson.Length)]}");

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
        await MatchItemsWithCatalogAsync(data.Items, docText, ct, log);

        return new AiParseResult
        {
            Data = data,
            RawResponseJson = responseJson,
            RawModelText = rawText,
            CleanedJson = cleanedJson
        };
    }

    private static string BuildUserPrompt(string docText, string productSection, string structuredDraft)
    {
        var builder = new StringBuilder();
        builder.AppendLine("以下是订单预解析草稿和订单全文。");
        builder.AppendLine("请提取 CustomerName、SaleDept、SaleManager、OrderDate、ContractAmount、FlexibleNet、Spec、BraidedSteelWireRope、UndergroundTransportMaxLength、Remark 和 Items。");
        builder.AppendLine("优先使用【预解析草稿】中的确定性结构；如草稿为空或字段缺失，再参考订单全文补全。");
        builder.AppendLine("Items 不能只看最下面的“产品及配件”列表，也要看全文中的真实物品行和预解析明细。");
        builder.AppendLine("重点：");
        builder.AppendLine("1. 第3部分会写主物品名称；主物品不一定是柔性网。紧随其后的每一行“规格：...”都是该主物品的一条物品明细。");
        builder.AppendLine("2. “整编钢丝绳：Φ21.5 钢丝绳 30米 金额：500”是物品。");
        builder.AppendLine("3. “整编钢丝绳要求：...”和“携带网卷...”“携带钢丝绳...”不是物品。");
        builder.AppendLine("4. 对“产品 型号 数量 单价”表格，严格按列读取，数量和单价不要读错。");
        builder.AppendLine("5. ItemName 必须保持订单原文中的叫法，不要自己改名，不要固定写成“柔性网”。");
        builder.AppendLine("6. 原文明写“金额”的物品，金额放 Amount，不要放 UnitPrice；表格里只有数量和单价时，Amount 按 Quantity * UnitPrice 计算。");
        builder.AppendLine("7. 如果只有一个物品缺少金额，它的 Amount 留空，由程序用订单总金额减去其它所有物品金额反推。");
        builder.AppendLine("8. 合同金额下面的“规格：xxx 单价 n”是第3部分主物品单价；第3部分“规格 xxx 宽 a 然后规格 yyy 宽 b”要拆成多条，金额=宽度*长度*单价。");
        builder.AppendLine("9. “混编：...”只作为要求说明，不解析成物品。");
        builder.AppendLine();

        if (!string.IsNullOrWhiteSpace(structuredDraft))
        {
            builder.AppendLine("【预解析草稿】");
            builder.AppendLine(structuredDraft);
            builder.AppendLine();
        }

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

    private static string BuildStructuredDraft(string docText)
    {
        var lines = GetNormalizedNonEmptyLines(docText);
        if (lines.Count == 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendDraftField(builder, "使用单位", lines.Select(line => ExtractAfterLabel(line, "使用单位")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)));
        AppendDraftField(builder, "合同金额", lines.Select(line => ExtractAfterLabel(line, "合同金额")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)));
        AppendDraftField(builder, "销售经理", lines.Select(line => ExtractAfterLabel(line, "销售经理")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)));
        AppendDraftField(builder, "销售部门", lines.Select(line => ExtractAfterLabel(line, "销售部门")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)));

        var sectionThreeLines = ExtractSectionLines(docText, "3");
        var sectionThreeSpecPrices = ExtractSectionSpecUnitPrices(lines);
        var mainItemName = FindSectionThreeMainItemName(sectionThreeLines);
        var mainItemLine = sectionThreeLines.FirstOrDefault(line => !string.IsNullOrWhiteSpace(TryExtractSectionThreeMainItemName(line)));
        var mainItemLength = ExtractSectionThreeMainLength(sectionThreeLines);
        if (!string.IsNullOrWhiteSpace(mainItemName))
        {
            builder.AppendLine();
            builder.AppendLine("【识别到的主产品】");
            builder.AppendLine($"物品名称：{mainItemName}");
            if (!string.IsNullOrWhiteSpace(mainItemLine))
            {
                builder.AppendLine($"主产品描述：{StripLeadingSerial(mainItemLine)}");
            }

            var totalLength = string.IsNullOrWhiteSpace(mainItemLength)
                ? ExtractNumberWithUnit(mainItemLine ?? string.Empty, @"(?:长|长度)\s*(?<value>\d+(?:\.\d+)?)\s*(?<unit>m|米)?")
                : mainItemLength + "米";
            AppendDraftField(builder, "总长度", totalLength);
        }

        var mainItems = sectionThreeLines
            .SelectMany(line => ParseSectionThreeSpecItems(line, mainItemName, mainItemLength, sectionThreeSpecPrices))
            .ToList();

        if (mainItems.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("【第3部分主物品明细】");
            for (var i = 0; i < mainItems.Count; i++)
            {
                AppendDraftItem(builder, i + 1, mainItems[i]);
            }
        }

        var ropeLine = lines.FirstOrDefault(line =>
            Regex.IsMatch(line, @"整编钢丝绳\s*[：:]") &&
            !line.Contains("要求", StringComparison.Ordinal));
        if (!string.IsNullOrWhiteSpace(ropeLine))
        {
            builder.AppendLine();
            builder.AppendLine("【整编用的钢丝绳】");
            builder.AppendLine(StripLeadingSerial(ropeLine));
        }

        var transportLine = lines.FirstOrDefault(line => line.Contains("井下运输最大尺寸", StringComparison.Ordinal));
        if (!string.IsNullOrWhiteSpace(transportLine))
        {
            builder.AppendLine();
            builder.AppendLine("【井下运输最大尺寸】");
            builder.AppendLine(StripLeadingSerial(transportLine));
            AppendDraftField(builder, "井下运输最大尺寸长", ExtractNumberWithUnit(transportLine, @"长\s*(?<value>\d+(?:\.\d+)?)\s*(?<unit>m|米)?"));
        }

        var tableItems = ExtractProductTableItemsFromLines(lines);
        if (tableItems.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("【产品及配件】");
            for (var i = 0; i < tableItems.Count; i++)
            {
                AppendDraftItem(builder, i + 1, tableItems[i]);
            }
        }

        return builder.ToString().Trim();
    }

    private static void AppendDraftField(StringBuilder builder, string name, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            builder.AppendLine($"{name}：{value.Trim()}");
        }
    }

    private static void AppendDraftItem(StringBuilder builder, int index, OrderItem item)
    {
        builder.AppendLine($"明细{index}：");
        AppendDraftField(builder, "  物品名称", item.ItemName);
        AppendDraftField(builder, "  规格", item.Spec);
        AppendDraftField(builder, "  型号/编码", item.ItemCode);
        AppendDraftField(builder, "  宽度", item.Width);
        AppendDraftField(builder, "  长度", item.Length);
        AppendDraftField(builder, "  长度分段", item.LengthSegments);
        AppendDraftField(builder, "  数量", item.Quantity);
        AppendDraftField(builder, "  单价", item.UnitPrice);
        AppendDraftField(builder, "  金额", item.Amount);
        AppendDraftField(builder, "  备注", item.ItemRemark);
    }

    private static List<string> GetNormalizedNonEmptyLines(string docText)
    {
        return docText
            .Replace("\r\n", "\n")
            .Split('\n')
            .Select(NormalizeSourceLine)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();
    }

    private static List<string> ExtractSectionLines(string docText, string sectionNumber)
    {
        var lines = GetNormalizedNonEmptyLines(docText);
        var sectionLines = new List<string>();
        var inSection = false;

        foreach (var line in lines)
        {
            var serial = ExtractLeadingSectionNumber(line);
            if (serial == sectionNumber)
            {
                inSection = true;
                sectionLines.Add(line);
                continue;
            }

            if (inSection && serial != null)
            {
                break;
            }

            if (inSection)
            {
                sectionLines.Add(line);
            }
        }

        return sectionLines;
    }

    private static string FindSectionThreeMainItemName(List<string> sectionThreeLines)
    {
        foreach (var line in sectionThreeLines)
        {
            var name = TryExtractSectionThreeMainItemName(line);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }
        }

        return string.Empty;
    }

    private static string TryExtractSectionThreeMainItemName(string line)
    {
        var content = StripLeadingSerial(line);
        if (string.IsNullOrWhiteSpace(content) ||
            Regex.IsMatch(content, @"^规格\s*[：:]") ||
            content.Contains("要求", StringComparison.Ordinal) ||
            content.Contains("产品及配件", StringComparison.Ordinal))
        {
            return string.Empty;
        }

        var sizeIndex = content.IndexOf("尺寸", StringComparison.Ordinal);
        if (sizeIndex > 0)
        {
            return CleanMainItemName(content[..sizeIndex]);
        }

        var lengthIndex = content.IndexOf("长度", StringComparison.Ordinal);
        if (lengthIndex > 0)
        {
            var prefix = content[..lengthIndex];
            var colonIndex = Math.Max(prefix.LastIndexOf('：'), prefix.LastIndexOf(':'));
            return CleanMainItemName(colonIndex >= 0 ? prefix[(colonIndex + 1)..] : prefix);
        }

        var match = Regex.Match(content, @"^(?<name>.+?)[：:]\s*(?:长|长度|宽|规格|型号)");
        return match.Success ? CleanMainItemName(match.Groups["name"].Value) : string.Empty;
    }

    private static string CleanMainItemName(string value)
    {
        var cleaned = value.Trim();
        cleaned = Regex.Replace(cleaned, @"^\d+\s*", string.Empty);
        cleaned = cleaned.Trim('：', ':', '，', ',', '。', ' ', '\t');
        return cleaned;
    }

    private static List<OrderItem> ParseSectionThreeSpecItems(
        string line,
        string itemName,
        string sectionLength,
        IReadOnlyDictionary<string, string> specUnitPrices)
    {
        var content = StripLeadingSerial(line);
        if (!Regex.IsMatch(content, @"^规格\s*[：:]"))
        {
            return [];
        }

        var items = new List<OrderItem>();
        foreach (Match match in Regex.Matches(
                     content,
                     @"规格\s*[：:]\s*(?<spec>.+?)(?:\s+宽\s*(?<width>\d+(?:\.\d+)?)\s*(?:m|米)?)?(?=\s*(?:然后)?规格\s*[：:]|金额\s*[：:]|$)",
                     RegexOptions.IgnoreCase))
        {
            var spec = Regex.Replace(match.Groups["spec"].Value, @"\s+", " ").Trim(' ', '，', ',', ';', '；');
            if (string.IsNullOrWhiteSpace(spec))
            {
                continue;
            }

            var width = match.Groups["width"].Success ? match.Groups["width"].Value.Trim() : string.Empty;
            var unitPrice = ResolveSpecUnitPrice(spec, specUnitPrices);
            var amount = ExtractNumber(content, @"金额\s*[：:]\s*(?<value>\d+(?:\.\d+)?)");
            var widthValue = ParseDecimal(width);
            var lengthValue = ParseDecimal(sectionLength);
            var unitPriceValue = ParseDecimal(unitPrice);
            if (widthValue.HasValue && lengthValue.HasValue && unitPriceValue.HasValue)
            {
                amount = (widthValue.Value * lengthValue.Value * unitPriceValue.Value).ToString("0.####", CultureInfo.InvariantCulture);
            }

            items.Add(new OrderItem
            {
                ItemName = itemName,
                Spec = spec,
                Width = width,
                Length = sectionLength,
                LengthSegments = ExtractNumber(content, @"长度分\s*(?<value>\d+(?:\.\d+)?)\s*段"),
                UnitPrice = unitPrice,
                Amount = amount,
                ItemRemark = match.Value.Trim()
            });
        }

        return items;
    }

    private static OrderItem? TryParseSectionThreeSpecLine(string line, string itemName)
    {
        return ParseSectionThreeSpecItems(line, itemName, string.Empty, new Dictionary<string, string>())
            .FirstOrDefault();
    }

    private static Dictionary<string, string> ExtractSectionSpecUnitPrices(IEnumerable<string> lines)
    {
        var prices = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var line in lines)
        {
            foreach (Match match in Regex.Matches(
                         line,
                         @"规格\s*[：:]\s*(?<spec>.+?)\s+单价\s*(?<price>\d+(?:\.\d+)?)",
                         RegexOptions.IgnoreCase))
            {
                var key = NormalizeSpecKey(match.Groups["spec"].Value);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    prices[key] = match.Groups["price"].Value.Trim();
                }
            }
        }

        return prices;
    }

    private static string ResolveSpecUnitPrice(string spec, IReadOnlyDictionary<string, string> specUnitPrices)
    {
        if (specUnitPrices.Count == 0)
        {
            return string.Empty;
        }

        var key = NormalizeSpecKey(spec);
        if (string.IsNullOrWhiteSpace(key))
        {
            return string.Empty;
        }

        if (specUnitPrices.TryGetValue(key, out var exactPrice))
        {
            return exactPrice;
        }

        foreach (var pair in specUnitPrices)
        {
            if (key.Contains(pair.Key, StringComparison.Ordinal) ||
                pair.Key.Contains(key, StringComparison.Ordinal))
            {
                return pair.Value;
            }
        }

        return string.Empty;
    }

    private static string NormalizeSpecKey(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.ToUpperInvariant()
            .Replace("Ｍ", "M", StringComparison.Ordinal)
            .Replace("—", "-", StringComparison.Ordinal)
            .Replace("－", "-", StringComparison.Ordinal);
        return Regex.Replace(normalized, @"[\s:：;；,，、。]+", string.Empty);
    }

    private static string ExtractSectionThreeMainLength(IEnumerable<string> sectionThreeLines)
    {
        foreach (var line in sectionThreeLines)
        {
            if (!line.Contains("尺寸", StringComparison.Ordinal) &&
                !line.Contains("长", StringComparison.Ordinal) &&
                !line.Contains("长度", StringComparison.Ordinal))
            {
                continue;
            }

            var length = ExtractNumber(line, @"(?:长|长度)\s*(?<value>\d+(?:\.\d+)?)\s*(?:m|米)?");
            if (!string.IsNullOrWhiteSpace(length))
            {
                return length;
            }
        }

        return string.Empty;
    }

    private static OrderItem? TryParseFlexibleNetSpecLine(string line) =>
        TryParseSectionThreeSpecLine(line, "柔性网");

    private static List<OrderItem> ExtractProductTableItemsFromLines(List<string> lines)
    {
        var start = lines.FindIndex(line => line.Contains("产品及配件", StringComparison.Ordinal));
        if (start < 0)
        {
            return [];
        }

        var tokens = CollectProductTableTokens(lines, start);

        var items = new List<OrderItem>();
        for (var i = 0; i + 3 < tokens.Count;)
        {
            var item = TryBuildTableItem(
                tokens[i],
                tokens[i + 1],
                tokens[i + 2],
                tokens[i + 3],
                string.Join(" ", tokens.Skip(i).Take(4)));

            if (item == null)
            {
                i++;
                continue;
            }

            items.Add(item);
            i += 4;
        }

        return items;
    }

    private static List<string> CollectProductTableTokens(List<string> lines, int start)
    {
        var tokens = new List<string>();
        for (var i = start + 1; i < lines.Count; i++)
        {
            var content = StripLeadingSerial(lines[i]);
            if (string.IsNullOrWhiteSpace(content))
            {
                continue;
            }

            if (content.Contains("其他要求", StringComparison.Ordinal) ||
                content.Contains("销售经理", StringComparison.Ordinal) ||
                content.Contains("销售部门", StringComparison.Ordinal) ||
                content.Contains("现场数据采集人员", StringComparison.Ordinal) ||
                content.Contains("为便于进一步沟通", StringComparison.Ordinal) ||
                Regex.IsMatch(content, @"^1[4-9]\b"))
            {
                break;
            }

            if (content is "产品" or "型号" or "数量" or "单价")
            {
                continue;
            }

            tokens.Add(content);
        }

        return tokens;
    }

    private static OrderData PostProcessOrderData(OrderData data, string docText, string productSection, Action<string>? log)
    {
        data.CustomerName = data.CustomerName?.Trim() ?? string.Empty;
        data.SaleDept = data.SaleDept?.Trim() ?? string.Empty;
        data.SaleManager = data.SaleManager?.Trim() ?? string.Empty;
        data.OrderDate = data.OrderDate?.Trim() ?? string.Empty;
        data.ContractAmount = data.ContractAmount?.Trim() ?? string.Empty;
        data.FlexibleNet = data.FlexibleNet?.Trim() ?? string.Empty;
        data.Spec = data.Spec?.Trim() ?? string.Empty;
        data.BraidedSteelWireRope = data.BraidedSteelWireRope?.Trim() ?? string.Empty;
        data.UndergroundTransportMaxLength = data.UndergroundTransportMaxLength?.Trim() ?? string.Empty;
        data.Remark = data.Remark?.Trim() ?? string.Empty;

        FillKeyFieldsFromSource(data, docText, log);

        var validProductLines = ExtractValidItemLines(docText, productSection);
        var filteredItems = new List<OrderItem>();

        foreach (var rawItem in data.Items ?? [])
        {
            var item = NormalizeItem(rawItem);
            if (!NormalizeNumericFields(item, log))
            {
                continue;
            }

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
        mergedItems = mergedItems
            .Where(item => NormalizeNumericFields(item, log))
            .ToList();
        mergedItems = RemoveUngroundedProductTableItems(mergedItems, docText, log);
        ApplySectionThreeMainItemQuantities(mergedItems, log);
        ApplyDefaultQuantity(mergedItems, log);
        mergedItems = DeduplicateMergedItems(mergedItems, log);

        ApplyAmountRules(mergedItems, log);
        FillSingleMissingAmountFromContract(mergedItems, data.ContractAmount, log);
        mergedItems = DeduplicateMergedItems(mergedItems, log);
        data.Items = mergedItems;
        return data;
    }

    private static void FillKeyFieldsFromSource(OrderData data, string docText, Action<string>? log)
    {
        var lines = docText
            .Replace("\r\n", "\n")
            .Split('\n')
            .Select(NormalizeSourceLine)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(data.CustomerName))
                data.CustomerName = ExtractAfterLabel(line, "使用单位");
            if (string.IsNullOrWhiteSpace(data.ContractAmount))
                data.ContractAmount = ExtractAfterLabel(line, "合同金额");
            if (string.IsNullOrWhiteSpace(data.SaleManager))
                data.SaleManager = ExtractAfterLabel(line, "销售经理");
            if (string.IsNullOrWhiteSpace(data.SaleDept))
                data.SaleDept = ExtractAfterLabel(line, "销售部门");

            if (string.IsNullOrWhiteSpace(data.FlexibleNet) &&
                line.Contains("柔性网尺寸", StringComparison.Ordinal))
            {
                data.FlexibleNet = StripLeadingSerial(line);
            }

            if (string.IsNullOrWhiteSpace(data.BraidedSteelWireRope) &&
                Regex.IsMatch(line, @"整编钢丝绳\s*[：:]") &&
                !line.Contains("要求", StringComparison.Ordinal))
            {
                data.BraidedSteelWireRope = StripLeadingSerial(line);
            }

            if (string.IsNullOrWhiteSpace(data.UndergroundTransportMaxLength) &&
                line.Contains("井下运输最大尺寸", StringComparison.Ordinal))
            {
                data.UndergroundTransportMaxLength = ExtractNumberWithUnit(line, @"长\s*(?<value>\d+(?:\.\d+)?)\s*(?<unit>m|米)?");
            }
        }

        if (string.IsNullOrWhiteSpace(data.Spec))
        {
            var specs = lines
                .Select(line => ExtractAfterLabel(StripLeadingSerial(line), "规格"))
                .Where(spec => !string.IsNullOrWhiteSpace(spec))
                .Select(CleanSpecSummary)
                .Where(spec => !string.IsNullOrWhiteSpace(spec))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            data.Spec = string.Join("；", specs);
        }

        log?.Invoke($"  关键字段：使用单位={data.CustomerName}，合同金额={data.ContractAmount}，规格={data.Spec}，销售经理={data.SaleManager}，销售部门={data.SaleDept}");
    }

    private static string ExtractAfterLabel(string line, string label)
    {
        var match = Regex.Match(line, $@"{Regex.Escape(label)}\s*[：:]\s*(?<value>.+)$");
        return match.Success ? match.Groups["value"].Value.Trim() : string.Empty;
    }

    private static string ExtractNumberWithUnit(string line, string pattern)
    {
        var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return string.Empty;
        }

        var value = match.Groups["value"].Value.Trim();
        var unit = match.Groups["unit"].Success ? match.Groups["unit"].Value.Trim() : "m";
        return $"{value}{unit}";
    }

    private static string ExtractValue(string line, string pattern)
    {
        var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase);
        return match.Success ? match.Groups["value"].Value.Trim() : string.Empty;
    }

    private static string CleanSpecSummary(string specLine)
    {
        var cleaned = Regex.Replace(specLine, @"金额\s*[：:].*$", string.Empty).Trim();
        cleaned = Regex.Replace(cleaned, @"\s+", " ");
        return cleaned.Trim('，', ',', ';', '；');
    }

    private static List<OrderItem> ExtractDeterministicItems(string docText, string productSection, Action<string>? log)
    {
        var items = new List<OrderItem>();
        var allLines = GetNormalizedNonEmptyLines(docText);
        var sectionThreeLines = ExtractSectionLines(docText, "3");
        var sectionThreeSpecPrices = ExtractSectionSpecUnitPrices(allLines);
        var sectionThreeMainItemName = FindSectionThreeMainItemName(sectionThreeLines);
        var sectionThreeLength = ExtractSectionThreeMainLength(sectionThreeLines);
        var sectionThreeItems = sectionThreeLines
            .SelectMany(line => ParseSectionThreeSpecItems(line, sectionThreeMainItemName, sectionThreeLength, sectionThreeSpecPrices))
            .Select(NormalizeItem)
            .ToList();
        items.AddRange(sectionThreeItems);

        var tableItems = ExtractProductTableItemsFromLines(allLines)
            .Select(NormalizeItem)
            .ToList();
        items.AddRange(tableItems);

        var sourceText = productSection;
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            if (items.Count > 0)
            {
                log?.Invoke($"  Source pre-parser parsed {items.Count} item(s).");
            }

            return items;
        }

        var lines = sourceText.Replace("\r\n", "\n").Split('\n');
        var currentItemName = string.Empty;
        var currentLength = string.Empty;
        var currentSection = string.Empty;
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

            var serial = ExtractLeadingSectionNumber(line);
            if (serial != null)
            {
                currentSection = serial;
            }

            if (currentSection == "3")
            {
                var detectedItemName = TryExtractSectionThreeMainItemName(contentLine);
                if (!string.IsNullOrWhiteSpace(detectedItemName))
                {
                    currentItemName = detectedItemName;
                    currentLength = ExtractNumber(contentLine, @"(?:长|长度)\s*(?<value>\d+(?:\.\d+)?)\s*(?:m|米)?");
                    continue;
                }
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

            var specItem = currentSection == "3"
                ? TryParseSpecItem(contentLine, currentItemName, currentLength)
                : null;
            if (specItem != null)
            {
                AddIfNotDuplicate(items, specItem);
                continue;
            }

            var ropeItem = TryParseSteelRopeItem(contentLine);
            if (ropeItem != null)
            {
                AddIfNotDuplicate(items, ropeItem);
                continue;
            }

            var transportFrameItem = TryParseTransportFrameItem(contentLine, nextLine);
            if (transportFrameItem != null)
            {
                AddIfNotDuplicate(items, transportFrameItem);
            }
        }

        if (items.Count > 0)
        {
            log?.Invoke($"  Source fallback parsed {items.Count} item(s).");
        }

        return items;
    }

    private static void AddIfNotDuplicate(List<OrderItem> items, OrderItem item)
    {
        var normalizedItem = NormalizeItem(item);
        var existingIndex = items.FindIndex(existing => AreLikelySameItem(existing, normalizedItem));
        if (existingIndex >= 0)
        {
            items[existingIndex] = PreferRicherItem(items[existingIndex], normalizedItem);
            return;
        }

        items.Add(normalizedItem);
    }

    private static bool AreLikelySameItem(OrderItem left, OrderItem right)
    {
        return NormalizeForMatch(left.ItemName) == NormalizeForMatch(right.ItemName) &&
               NormalizeForMatch(left.Spec) == NormalizeForMatch(right.Spec) &&
               NormalizeForMatch(left.Width) == NormalizeForMatch(right.Width) &&
               NormalizeForMatch(left.Length) == NormalizeForMatch(right.Length) &&
               NormalizeForMatch(left.LengthSegments) == NormalizeForMatch(right.LengthSegments) &&
               NormalizeForMatch(left.Quantity) == NormalizeForMatch(right.Quantity) &&
               NormalizeForMatch(left.UnitPrice) == NormalizeForMatch(right.UnitPrice) &&
               NormalizeForMatch(left.Amount) == NormalizeForMatch(right.Amount);
    }

    private static List<OrderItem> DeduplicateMergedItems(List<OrderItem> items, Action<string>? log)
    {
        var deduplicatedItems = new List<OrderItem>();
        var indexByKey = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var rawItem in items)
        {
            var item = NormalizeItem(rawItem);
            var key = BuildItemIdentityKey(item);
            if (string.IsNullOrWhiteSpace(key))
            {
                deduplicatedItems.Add(item);
                continue;
            }

            if (indexByKey.TryGetValue(key, out var existingIndex))
            {
                deduplicatedItems[existingIndex] = PreferRicherItem(deduplicatedItems[existingIndex], item);
                continue;
            }

            indexByKey[key] = deduplicatedItems.Count;
            deduplicatedItems.Add(item);
        }

        if (deduplicatedItems.Count != items.Count)
        {
            log?.Invoke($"  Removed duplicate merged items: {items.Count - deduplicatedItems.Count}");
        }

        return deduplicatedItems;
    }

    private static string BuildItemIdentityKey(OrderItem item)
    {
        var name = NormalizeForMatch(item.ItemName);
        var spec = NormalizeForMatch(item.Spec);
        var code = NormalizeForMatch(item.ItemCode);
        var width = NormalizeForMatch(item.Width);
        var lengthSegments = NormalizeForMatch(item.LengthSegments);
        var itemRemark = NormalizeForMatch(item.ItemRemark);
        var amount = NormalizeForMatch(item.Amount);

        if (!string.IsNullOrWhiteSpace(spec) &&
            !string.IsNullOrWhiteSpace(width) &&
            !string.IsNullOrWhiteSpace(lengthSegments))
        {
            var sourceLineAnchor = !string.IsNullOrWhiteSpace(itemRemark) ? itemRemark : amount;
            return $"spec-line|{name}|{spec}|{width}|{lengthSegments}|{sourceLineAnchor}";
        }

        if (IsLikelyProductTableItem(item))
        {
            return $"table|{code}|{name}|{spec}|{NormalizeForMatch(item.Quantity)}|{NormalizeForMatch(item.UnitPrice)}";
        }

        if (!string.IsNullOrWhiteSpace(itemRemark))
        {
            return $"remark|{name}|{spec}|{itemRemark}";
        }

        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(spec) && string.IsNullOrWhiteSpace(code))
        {
            return string.Empty;
        }

        return $"general|{code}|{name}|{spec}|{NormalizeForMatch(item.Length)}|{NormalizeForMatch(item.Quantity)}|{NormalizeForMatch(item.UnitPrice)}|{amount}";
    }

    private static OrderItem PreferRicherItem(OrderItem left, OrderItem right)
    {
        return ScoreItemCompleteness(right) >= ScoreItemCompleteness(left)
            ? MergeItems(left, right)
            : MergeItems(right, left);
    }

    private static int ScoreItemCompleteness(OrderItem item)
    {
        var score = 0;
        score += ScoreField(item.ItemCode);
        score += ScoreField(item.ItemName);
        score += ScoreField(item.Spec);
        score += ScoreField(item.Unit);
        score += ScoreField(item.Width, 2);
        score += ScoreField(item.Length, 2);
        score += ScoreField(item.LengthSegments, 2);
        score += ScoreField(item.Quantity, 2);
        score += ScoreField(item.UnitPrice, 2);
        score += ScoreField(item.Amount, 2);
        score += ScoreField(item.TaxRate);
        score += ScoreField(item.DeliveryDate);
        score += ScoreField(item.ItemRemark);
        return score;
    }

    private static int ScoreField(string? value, int weight = 1) =>
        string.IsNullOrWhiteSpace(value) ? 0 : weight;

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
                tableRowBuffer.RemoveRange(0, 4);
                continue;
            }

            tableRowBuffer.RemoveAt(0);
        }
    }

    private static OrderItem? TryParseSpecItem(string line, string currentItemName, string currentLength)
    {
        if (line.Contains("然后规格", StringComparison.Ordinal))
        {
            return null;
        }

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

    private static bool NormalizeNumericFields(OrderItem item, Action<string>? log)
    {
        var hasInvalidQuantity = !string.IsNullOrWhiteSpace(item.Quantity) && ParseDecimal(item.Quantity) == null;
        var hasInvalidUnitPrice = !string.IsNullOrWhiteSpace(item.UnitPrice) && ParseDecimal(item.UnitPrice) == null;

        if (!hasInvalidQuantity && !hasInvalidUnitPrice)
        {
            return true;
        }

        var originalQuantity = item.Quantity;
        var originalUnitPrice = item.UnitPrice;

        if (hasInvalidQuantity)
        {
            item.Quantity = string.Empty;
        }

        if (hasInvalidUnitPrice)
        {
            item.UnitPrice = string.Empty;
        }

        if (string.IsNullOrWhiteSpace(item.Amount) && !IsSectionThreeMainItem(item))
        {
            log?.Invoke($"  已丢弃数量/单价错位的物品: 名称={item.ItemName}, 规格={item.Spec}, 原数量={originalQuantity}, 原单价={originalUnitPrice}");
            return false;
        }

        log?.Invoke($"  已清理数量/单价错位字段: 名称={item.ItemName}, 规格={item.Spec}, 原数量={originalQuantity}, 原单价={originalUnitPrice}");
        return true;
    }

    private static List<OrderItem> RemoveUngroundedProductTableItems(List<OrderItem> items, string docText, Action<string>? log)
    {
        var lines = GetNormalizedNonEmptyLines(docText);
        var productTableStart = lines.FindIndex(line => line.Contains("产品及配件", StringComparison.Ordinal));
        if (productTableStart < 0)
        {
            return items;
        }

        var productTokens = CollectProductTableTokens(lines, productTableStart);
        if (productTokens.Count < 4)
        {
            return items;
        }

        var filtered = new List<OrderItem>();
        foreach (var item in items)
        {
            if (IsLikelyProductTableItem(item) && !IsProductTableItemGrounded(item, productTokens))
            {
                log?.Invoke($"  已丢弃未在产品表中逐列命中的物品: 名称={item.ItemName}, 规格={item.Spec}, 数量={item.Quantity}, 单价={item.UnitPrice}");
                continue;
            }

            filtered.Add(item);
        }

        return filtered;
    }

    private static bool IsProductTableItemGrounded(OrderItem item, List<string> productTokens)
    {
        for (var i = 0; i + 3 < productTokens.Count; i++)
        {
            var quantity = ParseDecimal(productTokens[i + 2]);
            var unitPrice = ParseDecimal(productTokens[i + 3]);
            var itemQuantity = ParseDecimal(item.Quantity);
            var itemUnitPrice = ParseDecimal(item.UnitPrice);
            if (!quantity.HasValue ||
                !unitPrice.HasValue ||
                !itemQuantity.HasValue ||
                !itemUnitPrice.HasValue ||
                quantity.Value != itemQuantity.Value ||
                unitPrice.Value != itemUnitPrice.Value)
            {
                continue;
            }

            if ((ProductTextMatches(item.ItemName, productTokens[i]) &&
                 ProductTextMatches(item.Spec, productTokens[i + 1])) ||
                ProductTextMatches(item.ItemName, productTokens[i]) ||
                ProductTextMatches($"{item.ItemName} {item.Spec}", $"{productTokens[i]} {productTokens[i + 1]}"))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ProductTextMatches(string itemValue, string sourceValue)
    {
        var itemText = NormalizeForMatch(itemValue);
        var sourceText = NormalizeForMatch(sourceValue);
        if (string.IsNullOrWhiteSpace(itemText) || string.IsNullOrWhiteSpace(sourceText))
        {
            return false;
        }

        return itemText == sourceText ||
               itemText.Contains(sourceText, StringComparison.Ordinal) ||
               sourceText.Contains(itemText, StringComparison.Ordinal);
    }

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

    private async Task MatchItemsWithCatalogAsync(List<OrderItem> items, string docText, CancellationToken ct, Action<string>? log)
    {
        var catalog = ItemCatalog.Load();
        var unmatchedItems = new List<(int Index, OrderItem Item)>();

        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            if (IsEmptyItem(item))
            {
                continue;
            }

            var match = catalog.FindMatch(item);
            if (match == null)
            {
                unmatchedItems.Add((i, item));
                continue;
            }

            ApplyCatalogMatch(item, match);
            log?.Invoke($"  物品匹配成功: {item.ItemName} / {item.Spec} / 编码={item.ItemCode}");
        }

        if (unmatchedItems.Count > 0)
        {
            log?.Invoke($"  本地规则有 {unmatchedItems.Count} 个物品未匹配，开始调用 AI 按物品分类信息表匹配...");
            await MatchUnmatchedItemsWithAiAsync(items, unmatchedItems, catalog, docText, ct, log);
        }

        var errors = items
            .Where(item => !IsEmptyItem(item) &&
                           (string.IsNullOrWhiteSpace(item.ItemCode) ||
                            catalog.FindByCode(item.ItemCode) == null))
            .Select(item => $"名称={item.ItemName}，规格={item.Spec}")
            .ToList();

        if (errors.Count > 0)
        {
            var message = "物品分类信息表中未匹配到以下物品，已停止解析：" + Environment.NewLine +
                          string.Join(Environment.NewLine, errors.Select(error => "  - " + error));
            throw new OrderItemCatalogMatchException(message);
        }
    }

    private async Task MatchUnmatchedItemsWithAiAsync(
        List<OrderItem> items,
        List<(int Index, OrderItem Item)> unmatchedItems,
        ItemCatalog catalog,
        string docText,
        CancellationToken ct,
        Action<string>? log)
    {
        var sourceText = ExtractRelevantItemSections(docText);
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            sourceText = docText;
        }

        if (sourceText.Length > CatalogMatchSourceMaxChars)
        {
            sourceText = sourceText[..CatalogMatchSourceMaxChars];
        }

        var requestItems = unmatchedItems
            .Select(item => new
            {
                item.Index,
                item.Item.ItemName,
                item.Item.Spec,
                item.Item.Quantity,
                item.Item.UnitPrice,
                item.Item.Amount,
                item.Item.ItemRemark
            })
            .ToList();

        var catalogRows = unmatchedItems
            .Select(unmatched => new
            {
                unmatched.Index,
                Candidates = catalog.FindAiMatchCandidates(unmatched.Item, sourceText, CatalogMatchCandidatesPerItem)
                    .Select(entry => new
                    {
                        entry.ItemCode,
                        entry.ItemName,
                        entry.Spec,
                        entry.Unit,
                        entry.CategoryName
                    })
                    .ToList()
            })
            .ToList();

        var prompt = new StringBuilder();
        prompt.AppendLine("请把【待匹配物品】匹配到【候选物品分类信息】中的唯一物品。");
        prompt.AppendLine("要求：");
        prompt.AppendLine("1. 物品名和规格可能被 Word/AI 解析错位、缺少空格、缺少数字，必须结合【订单原文片段】理解。");
        prompt.AppendLine("2. 每个 Index 只能从对应 Index 的 Candidates 中选择 ItemCode，不能编造。");
        prompt.AppendLine("3. 匹配成功后，程序会使用该 ItemCode 对应的 CSV 物品名称、规格型号、计量单位；不要返回自己改写的名称或规格。");
        prompt.AppendLine("4. 如果无法确定，ItemCode 填空字符串。");
        prompt.AppendLine("5. 只返回 JSON，不要解释。格式：{\"Matches\":[{\"Index\":0,\"ItemCode\":\"07020024\"}]}");
        prompt.AppendLine();
        prompt.AppendLine("【订单原文片段】");
        prompt.AppendLine(sourceText);
        prompt.AppendLine();
        prompt.AppendLine("【待匹配物品】");
        prompt.AppendLine(JsonSerializer.Serialize(requestItems, ParsedCompactJsonOptions));
        prompt.AppendLine();
        prompt.AppendLine("【候选物品分类信息】");
        prompt.AppendLine(JsonSerializer.Serialize(catalogRows, ParsedCompactJsonOptions));
        log?.Invoke($"  AI 物品匹配请求大小: 未匹配 {unmatchedItems.Count} 个，候选 {catalogRows.Sum(row => row.Candidates.Count)} 行，prompt约 {prompt.Length:N0} 字，max_tokens={CatalogMatchMaxTokens}");

        var url = baseUrl.TrimEnd('/') + "/v1/messages";
        var body = new
        {
            model,
            max_tokens = CatalogMatchMaxTokens,
            stream = false,
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = prompt.ToString()
                }
            }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

        HttpResponseMessage response;
        var stopwatch = Stopwatch.StartNew();
        try
        {
            response = await Http.SendAsync(request, ct);
        }
        catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
        {
            throw new TimeoutException("AI 物品匹配请求超过 10 分钟仍未返回。请稍后重试，或检查网络和模型响应速度。", ex);
        }

        var responseJson = await response.Content.ReadAsStringAsync(ct);
        stopwatch.Stop();
        log?.Invoke($"  AI 物品匹配接口耗时 {stopwatch.Elapsed.TotalSeconds:F1} 秒，状态 {(int)response.StatusCode}");
        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"AI 物品匹配接口返回 {(int)response.StatusCode}: {responseJson}");
        }

        var rawText = ExtractModelText(responseJson);
        var cleanedJson = ExtractJsonObject(rawText);
        var aiResponse = JsonSerializer.Deserialize<CatalogAiResponse>(cleanedJson, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        foreach (var aiMatch in aiResponse?.Matches ?? [])
        {
            if (string.IsNullOrWhiteSpace(aiMatch.ItemCode) ||
                aiMatch.Index < 0 ||
                aiMatch.Index >= items.Count)
            {
                continue;
            }

            var catalogEntry = catalog.FindByCode(aiMatch.ItemCode);
            if (catalogEntry == null)
            {
                continue;
            }

            ApplyCatalogMatch(items[aiMatch.Index], catalogEntry);
            log?.Invoke($"  AI 物品匹配成功: 原名称={requestItems.FirstOrDefault(item => item.Index == aiMatch.Index)?.ItemName}, 标准={catalogEntry.ItemName} / {catalogEntry.Spec} / 编码={catalogEntry.ItemCode}");
        }
    }

    private static readonly JsonSerializerOptions ParsedCompactJsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static void ApplyCatalogMatch(OrderItem item, ItemCatalogEntry match)
    {
        item.ItemCode = match.ItemCode;
        item.ItemName = match.ItemName;
        item.Spec = match.Spec;
        item.Unit = match.Unit;
    }

    private static void ApplyAmountRules(List<OrderItem> items, Action<string>? log)
    {
        foreach (var item in items)
        {
            var amount = ParseDecimal(item.Amount);
            if (amount.HasValue && amount.Value > 0)
            {
                continue;
            }

            var quantity = ParseDecimal(item.Quantity);
            var unitPrice = ParseDecimal(item.UnitPrice);
            if (!quantity.HasValue || !unitPrice.HasValue)
            {
                if (!string.IsNullOrWhiteSpace(item.Amount) && (!amount.HasValue || amount.Value <= 0))
                {
                    item.Amount = string.Empty;
                }

                continue;
            }

            var calculatedAmount = quantity.Value * unitPrice.Value;
            if (calculatedAmount <= 0)
            {
                item.Amount = string.Empty;
                continue;
            }

            item.Amount = calculatedAmount.ToString("0.####", CultureInfo.InvariantCulture);
            log?.Invoke($"  已按数量*单价计算金额: 名称={item.ItemName}, 规格={item.Spec}, 金额={item.Amount}");
        }
    }

    private static void FillSingleMissingAmountFromContract(List<OrderItem> items, string contractAmount, Action<string>? log)
    {
        var contractTotal = ParseDecimal(contractAmount);
        if (contractTotal == null || items.Count == 0)
        {
            return;
        }

        var missingAmountItems = items
            .Where(item => IsMissingAmountValue(item.Amount))
            .ToList();

        if (missingAmountItems.Count != 1)
        {
            return;
        }

        decimal knownTotal = 0;
        foreach (var item in items)
        {
            if (ReferenceEquals(item, missingAmountItems[0]))
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

        var missingItem = missingAmountItems[0];
        missingItem.Amount = remaining.ToString("0.####", CultureInfo.InvariantCulture);

        log?.Invoke($"  已按合同总价反推唯一缺失物品金额: 名称={missingItem.ItemName}, 金额={missingItem.Amount}");
    }

    private static bool IsMissingAmountValue(string? value) =>
        string.IsNullOrWhiteSpace(value) || ParseDecimal(value) is not { } amount || amount <= 0;

    private static decimal? ResolveLineTotal(OrderItem item)
    {
        var amount = ParseDecimal(item.Amount);
        if (amount.HasValue)
        {
            return amount.Value;
        }

        var quantity = ParseDecimal(item.Quantity);
        var unitPrice = ParseDecimal(item.UnitPrice);
        if (quantity.HasValue && unitPrice.HasValue)
        {
            return quantity.Value * unitPrice.Value;
        }

        return null;
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
        var targets = new HashSet<string>(StringComparer.Ordinal) { "1", "3", "4", "6", "13", "14" };
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

public sealed class OrderItemCatalogMatchException(string message) : Exception(message);

internal sealed class CatalogAiResponse
{
    public List<CatalogAiMatch> Matches { get; set; } = [];
}

internal sealed class CatalogAiMatch
{
    public int Index { get; set; }
    public string ItemCode { get; set; } = "";
}

internal sealed class ItemCatalog
{
    private const string CatalogFileName = "物品分类信息表.csv";
    private static readonly Lazy<ItemCatalog> Instance = new(LoadCore);
    private readonly List<ItemCatalogEntry> _entries;

    private ItemCatalog(List<ItemCatalogEntry> entries)
    {
        _entries = entries;
    }

    public static ItemCatalog Load() => Instance.Value;

    public IReadOnlyList<ItemCatalogEntry> Entries => _entries;

    public ItemCatalogEntry? FindByCode(string itemCode)
    {
        var normalizedCode = NormalizeCode(itemCode);
        return string.IsNullOrWhiteSpace(normalizedCode)
            ? null
            : _entries.FirstOrDefault(entry => NormalizeCode(entry.ItemCode) == normalizedCode);
    }

    public ItemCatalogEntry? FindMatch(OrderItem item)
    {
        var itemCode = NormalizeCode(item.ItemCode);
        if (!string.IsNullOrWhiteSpace(itemCode))
        {
            var codeMatch = _entries.FirstOrDefault(entry => NormalizeCode(entry.ItemCode) == itemCode);
            if (codeMatch != null)
            {
                return codeMatch;
            }
        }

        var name = NormalizeForCatalogMatch(item.ItemName);
        var spec = NormalizeForCatalogMatch(item.Spec);

        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(spec))
        {
            return null;
        }

        var candidates = string.IsNullOrWhiteSpace(spec)
            ? _entries.Where(entry => NamesMatch(name, entry)).ToList()
            : _entries.Where(entry => entry.NormalizedSpec == spec).ToList();

        if (candidates.Count == 0)
        {
            return null;
        }

        var exactNameMatches = candidates
            .Where(entry => entry.NormalizedItemName == name || entry.NormalizedCategoryName == name)
            .ToList();
        if (exactNameMatches.Count == 1)
        {
            return exactNameMatches[0];
        }

        var fuzzyNameMatches = candidates
            .Where(entry => NamesMatch(name, entry))
            .ToList();
        if (fuzzyNameMatches.Count == 1)
        {
            return fuzzyNameMatches[0];
        }

        if (!string.IsNullOrWhiteSpace(spec) && candidates.Count == 1)
        {
            return candidates[0];
        }

        return null;
    }

    public IReadOnlyList<ItemCatalogEntry> FindAiMatchCandidates(OrderItem item, string sourceText, int maxCandidates)
    {
        if (maxCandidates <= 0)
        {
            return [];
        }

        var scored = _entries
            .Select(entry => new
            {
                Entry = entry,
                Score = ScoreAiCandidate(item, sourceText, entry)
            })
            .Where(candidate => candidate.Score > 0)
            .OrderByDescending(candidate => candidate.Score)
            .ThenBy(candidate => candidate.Entry.ItemCode, StringComparer.Ordinal)
            .Take(maxCandidates)
            .Select(candidate => candidate.Entry)
            .ToList();

        return scored.Count > 0
            ? scored
            : _entries.Take(maxCandidates).ToList();
    }

    private static ItemCatalog LoadCore()
    {
        var path = ResolveCatalogPath();
        var lines = ReadCatalogText(path)
            .Replace("\r\n", "\n")
            .Split('\n')
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();

        if (lines.Count == 0)
        {
            throw new InvalidOperationException($"物品分类信息表为空：{path}");
        }

        var header = ParseCsvLine(lines[0]);
        var categoryCodeIndex = FindColumnIndex(header, "物品分类编码");
        var categoryNameIndex = FindColumnIndex(header, "物品分类名称");
        var itemCodeIndex = FindColumnIndex(header, "物品编码");
        var itemNameIndex = FindColumnIndex(header, "物品名称");
        var specIndex = FindColumnIndex(header, "规格型号");
        var unitIndex = FindColumnIndex(header, "计量单位");

        var entries = new List<ItemCatalogEntry>();
        foreach (var line in lines.Skip(1))
        {
            var columns = ParseCsvLine(line);
            var itemCode = GetColumn(columns, itemCodeIndex);
            var itemName = GetColumn(columns, itemNameIndex);
            if (string.IsNullOrWhiteSpace(itemCode) || string.IsNullOrWhiteSpace(itemName))
            {
                continue;
            }

            entries.Add(new ItemCatalogEntry(
                GetColumn(columns, categoryCodeIndex),
                GetColumn(columns, categoryNameIndex),
                itemCode,
                itemName,
                GetColumn(columns, specIndex),
                GetColumn(columns, unitIndex)));
        }

        if (entries.Count == 0)
        {
            throw new InvalidOperationException($"物品分类信息表没有有效物品行：{path}");
        }

        return new ItemCatalog(entries);
    }

    private static string ResolveCatalogPath()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "DataFile", CatalogFileName),
            Path.Combine(Directory.GetCurrentDirectory(), "DataFile", CatalogFileName),
            Path.Combine(Directory.GetCurrentDirectory(), "AIWorkAssistant", "DataFile", CatalogFileName)
        };

        var path = candidates.FirstOrDefault(File.Exists);
        if (!string.IsNullOrWhiteSpace(path))
        {
            return path;
        }

        throw new FileNotFoundException($"未找到物品分类信息表：{CatalogFileName}");
    }

    private static string ReadCatalogText(string path)
    {
        var bytes = File.ReadAllBytes(path);
        try
        {
            var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
            var text = utf8.GetString(bytes);
            if (text.Contains("物品编码", StringComparison.Ordinal))
            {
                return text;
            }
        }
        catch (DecoderFallbackException)
        {
            // Fall back to the Windows Chinese encoding used by the current CSV.
        }

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding("GB18030").GetString(bytes);
    }

    private static int FindColumnIndex(List<string> header, string columnName)
    {
        var index = header.FindIndex(column => column.Trim() == columnName);
        if (index < 0)
        {
            throw new InvalidOperationException($"物品分类信息表缺少列：{columnName}");
        }

        return index;
    }

    private static string GetColumn(List<string> columns, int index) =>
        index >= 0 && index < columns.Count ? columns[index].Trim() : string.Empty;

    private static List<string> ParseCsvLine(string line)
    {
        var values = new List<string>();
        var builder = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    builder.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }

                continue;
            }

            if (ch == ',' && !inQuotes)
            {
                values.Add(builder.ToString());
                builder.Clear();
                continue;
            }

            builder.Append(ch);
        }

        values.Add(builder.ToString());
        return values;
    }

    private static bool NamesMatch(string normalizedName, ItemCatalogEntry entry)
    {
        if (string.IsNullOrWhiteSpace(normalizedName))
        {
            return false;
        }

        return entry.NormalizedItemName == normalizedName ||
               entry.NormalizedCategoryName == normalizedName ||
               entry.NormalizedItemName.Contains(normalizedName, StringComparison.Ordinal) ||
               normalizedName.Contains(entry.NormalizedItemName, StringComparison.Ordinal) ||
               entry.NormalizedCategoryName.Contains(normalizedName, StringComparison.Ordinal) ||
               normalizedName.Contains(entry.NormalizedCategoryName, StringComparison.Ordinal);
    }

    private static int ScoreAiCandidate(OrderItem item, string sourceText, ItemCatalogEntry entry)
    {
        var name = NormalizeForCatalogMatch(item.ItemName);
        var spec = NormalizeForCatalogMatch(item.Spec);
        var itemText = $"{item.ItemName} {item.Spec} {item.ItemRemark}";
        var query = NormalizeForCatalogMatch(itemText);
        var source = NormalizeForCatalogMatch(sourceText);
        var score = 0;

        if (!string.IsNullOrWhiteSpace(name))
        {
            if (entry.NormalizedItemName == name)
            {
                score += 200;
            }
            else if (NamesMatch(name, entry))
            {
                score += 120;
            }
        }

        if (!string.IsNullOrWhiteSpace(spec) && spec.Length > 1)
        {
            if (entry.NormalizedSpec == spec)
            {
                score += 100;
            }
            else if (!string.IsNullOrWhiteSpace(entry.NormalizedSpec) &&
                     (entry.NormalizedSpec.Contains(spec, StringComparison.Ordinal) ||
                      spec.Contains(entry.NormalizedSpec, StringComparison.Ordinal)))
            {
                score += 45;
            }
        }

        if (!string.IsNullOrWhiteSpace(query))
        {
            if (!string.IsNullOrWhiteSpace(entry.NormalizedItemName) &&
                query.Contains(entry.NormalizedItemName, StringComparison.Ordinal))
            {
                score += 80;
            }

            if (!string.IsNullOrWhiteSpace(entry.NormalizedSpec) &&
                entry.NormalizedSpec.Length > 1 &&
                query.Contains(entry.NormalizedSpec, StringComparison.Ordinal))
            {
                score += 55;
            }
        }

        foreach (var token in ExtractCatalogMatchTokens(itemText))
        {
            if (entry.NormalizedItemName.Contains(token, StringComparison.Ordinal) ||
                entry.NormalizedCategoryName.Contains(token, StringComparison.Ordinal))
            {
                score += 35;
            }

            if (!string.IsNullOrWhiteSpace(entry.NormalizedSpec) &&
                entry.NormalizedSpec.Contains(token, StringComparison.Ordinal))
            {
                score += 25;
            }
        }

        foreach (var token in ExtractCatalogMatchTokens(sourceText).Take(80))
        {
            if (entry.NormalizedItemName.Contains(token, StringComparison.Ordinal) ||
                entry.NormalizedSpec.Contains(token, StringComparison.Ordinal))
            {
                score += 4;
            }
        }

        if (!string.IsNullOrWhiteSpace(source))
        {
            if (!string.IsNullOrWhiteSpace(entry.NormalizedItemName) &&
                source.Contains(entry.NormalizedItemName, StringComparison.Ordinal))
            {
                score += 30;
            }

            if (!string.IsNullOrWhiteSpace(entry.NormalizedSpec) &&
                entry.NormalizedSpec.Length > 1 &&
                source.Contains(entry.NormalizedSpec, StringComparison.Ordinal))
            {
                score += 15;
            }
        }

        return score;
    }

    private static IEnumerable<string> ExtractCatalogMatchTokens(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return [];
        }

        var tokens = new List<string>();
        foreach (Match match in Regex.Matches(
                     value,
                     @"[ΦφФф]\s*\d+(?:\.\d+)?|[A-Za-z]*\d+(?:\.\d+)?(?:\s*[*×xX]\s*\d+(?:\.\d+)?)*(?:\s*(?:mm|毫米|m|米|t|T))?",
                     RegexOptions.IgnoreCase))
        {
            AddToken(match.Value);
        }

        var chinese = new StringBuilder();
        foreach (var ch in value)
        {
            if (IsCjk(ch))
            {
                chinese.Append(ch);
                continue;
            }

            FlushChineseToken();
        }

        FlushChineseToken();

        return tokens
            .Select(NormalizeForCatalogMatch)
            .Where(token => token.Length >= 2 && token != "M")
            .Distinct(StringComparer.Ordinal)
            .ToList();

        void AddToken(string token)
        {
            if (!string.IsNullOrWhiteSpace(token))
            {
                tokens.Add(token);
            }
        }

        void FlushChineseToken()
        {
            if (chinese.Length >= 2)
            {
                tokens.Add(chinese.ToString());
            }

            chinese.Clear();
        }
    }

    private static bool IsCjk(char ch) =>
        ch is >= '\u3400' and <= '\u9fff';

    private static string NormalizeCode(string? value) =>
        string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static string NormalizeForCatalogMatch(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Trim().ToUpperInvariant()
            .Replace('φ', 'Φ')
            .Replace('Ф', 'Φ')
            .Replace('ф', 'Φ')
            .Replace("米", "M", StringComparison.Ordinal)
            .Replace("Ｍ", "M", StringComparison.Ordinal);

        normalized = Regex.Replace(normalized, @"[\s\r\n\t:：;；,，、。""'`·\[\]【】\(\)（）<>《》]+", string.Empty);
        return normalized;
    }
}

internal sealed class ItemCatalogEntry
{
    public ItemCatalogEntry(
        string categoryCode,
        string categoryName,
        string itemCode,
        string itemName,
        string spec,
        string unit)
    {
        CategoryCode = categoryCode.Trim();
        CategoryName = categoryName.Trim();
        ItemCode = itemCode.Trim();
        ItemName = itemName.Trim();
        Spec = spec.Trim();
        Unit = unit.Trim();
        NormalizedCategoryName = NormalizeForCatalogMatch(CategoryName);
        NormalizedItemName = NormalizeForCatalogMatch(ItemName);
        NormalizedSpec = NormalizeForCatalogMatch(Spec);
    }

    public string CategoryCode { get; }
    public string CategoryName { get; }
    public string ItemCode { get; }
    public string ItemName { get; }
    public string Spec { get; }
    public string Unit { get; }
    public string NormalizedCategoryName { get; }
    public string NormalizedItemName { get; }
    public string NormalizedSpec { get; }

    private static string NormalizeForCatalogMatch(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Trim().ToUpperInvariant()
            .Replace('φ', 'Φ')
            .Replace('Ф', 'Φ')
            .Replace('ф', 'Φ')
            .Replace("米", "M", StringComparison.Ordinal)
            .Replace("Ｍ", "M", StringComparison.Ordinal);

        normalized = Regex.Replace(normalized, @"[\s\r\n\t:：;；,，、。""'`·\[\]【】\(\)（）<>《》]+", string.Empty);
        return normalized;
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

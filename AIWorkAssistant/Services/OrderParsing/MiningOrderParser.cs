using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace AIWorkAssistant.Services.OrderParsing;

public static partial class DocumentTextExtractor
{
    public static async Task<string> ExtractAsync(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("文件不存在", path);

        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is ".txt" or ".text")
        {
            return Normalize(await File.ReadAllTextAsync(path, Encoding.UTF8));
        }

        var textutil = await TryRunAsync("textutil", $"-convert txt -stdout {Quote(path)}");
        if (!string.IsNullOrWhiteSpace(textutil)) return Normalize(textutil);

        var tempDir = Path.Combine(Path.GetTempPath(), "order-parser-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            await TryRunAsync("libreoffice", $"--headless --convert-to txt --outdir {Quote(tempDir)} {Quote(path)}");
            var txt = Directory.GetFiles(tempDir, "*.txt").FirstOrDefault();
            if (txt != null)
            {
                return Normalize(await File.ReadAllTextAsync(txt, Encoding.UTF8));
            }
        }
        finally
        {
            try { Directory.Delete(tempDir, true); } catch { /* ignore */ }
        }

        throw new InvalidOperationException("无法抽取 Word 文本。请安装 LibreOffice，或先把 Word 另存为 txt/docx。Mac 上通常可直接使用系统 textutil。");
    }

    private static async Task<string?> TryRunAsync(string fileName, string arguments)
    {
        try
        {
            var psi = new ProcessStartInfo(fileName, arguments)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            using var process = Process.Start(psi);
            if (process == null) return null;
            var stdout = await process.StandardOutput.ReadToEndAsync();
            await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();
            return process.ExitCode == 0 ? stdout : null;
        }
        catch
        {
            return null;
        }
    }

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

    public static string Normalize(string text)
    {
        text = text.Replace('\u3000', ' ')
            .Replace("\a", "\n")
            .Replace("\r\n", "\n")
            .Replace("\r", "\n");
        text = Regex.Replace(text, "[ \\t]+", " ");
        text = Regex.Replace(text, "\\n\\s+", "\n");
        text = Regex.Replace(text, "\\s+\\n", "\n");
        text = Regex.Replace(text, "\\n{3,}", "\n\n");
        return text.Trim();
    }
}

public static partial class OrderParser
{
    public static ParsedOrder Parse(string rawText)
    {
        var text = DocumentTextExtractor.Normalize(rawText);
        var sections = SplitSections(text);
        var accessorySection = sections.Values.FirstOrDefault(x => x.Contains("产品及配件", StringComparison.Ordinal)) ??
                               sections.GetValueOrDefault("11") ??
                               sections.GetValueOrDefault("13") ??
                               string.Empty;

        var order = new ParsedOrder
        {
            BasicInfo = ParseBasicInfo(sections.GetValueOrDefault("1") ?? string.Empty),
            Nets = ParseSection3(sections.GetValueOrDefault("3") ?? string.Empty, out var carryItems),
            CarryItems = carryItems,
            UndergroundTransportSizes = ParseUndergroundTransport(sections.GetValueOrDefault("6") ?? string.Empty),
            Accessories = ParseAccessories(accessorySection),
            Sales = ParseSalesFooter(text),
            RawSectionKeys = sections.Keys.OrderBy(x => int.TryParse(x, out var n) ? n : 999).ToList(),
        };

        order.Validation = Validate(order);
        return order;
    }

    private static Dictionary<string, string> SplitSections(string text)
    {
        var result = new Dictionary<string, List<string>>();
        string? current = null;
        var lastNo = 0;

        foreach (var raw in text.Split('\n'))
        {
            var line = raw.Trim();
            if (line.Length == 0) continue;

            if (Regex.IsMatch(line, "^\\d{1,2}$") && int.TryParse(line, out var no))
            {
                // 只接受模板范围内递增序号，避免配件数量 15/40/50 被误判成章节。
                if (no >= 1 && no <= 14 && no > lastNo)
                {
                    current = no.ToString();
                    lastNo = no;
                    result.TryAdd(current, []);
                    continue;
                }
            }

            if (current != null) result[current].Add(line);
        }

        return result.ToDictionary(kv => kv.Key, kv => string.Join('\n', kv.Value).Trim());
    }

    private static BasicInfo ParseBasicInfo(string section)
    {
        var info = new BasicInfo
        {
            Customer = MatchValue(section, "使用单位[:：]\\s*(.*?)\\s*[；;]"),
            SalesType = MatchValue(section, "销售类型\\s*[:：]\\s*([^\\n；;]+)"),
            ContractTotalAmount = MatchDecimal(section, "合同金额人民币\\s*[:：]\\s*([\\d,.]+)"),
            ExchangeRate = MatchDecimal(section, "汇率\\s*[:：]\\s*([\\d.]+)"),
            DeliveryDate = ParseChineseDate(section)
        };

        foreach (var part in Regex.Split(section, "[；;]\\s*"))
        {
            var m = Regex.Match(part,
                @"(?<name>柔性网|硬质网|一般钢丝绳)\s*(?:规格\s*)?(?<spec>[Φφ]?\s*[\d.*\-]+)?\s*(?:(?:单价\s*(?<unitPrice>[\d.]+)\s*元)|(?:金额\s*(?<amount>[\d.]+)\s*元)|(?<pricingMethod>倒算))");
            if (!m.Success) continue;

            info.PriceTerms.Add(new PriceTerm
            {
                Name = Clean(m.Groups["name"].Value),
                Spec = NormalizeSpec(m.Groups["spec"].Success ? m.Groups["spec"].Value : null),
                UnitPrice = ToDecimal(m.Groups["unitPrice"].Success ? m.Groups["unitPrice"].Value : null),
                Amount = ToDecimal(m.Groups["amount"].Success ? m.Groups["amount"].Value : null),
                PricingMethod = m.Groups["pricingMethod"].Success ? m.Groups["pricingMethod"].Value : null,
            });
        }

        return info;
    }

    private static List<NetItem> ParseSection3(string section, out List<CarryItem> carryItems)
    {
        var blocks = SplitNetBlocks(section, out var carryText);
        carryItems = ParseCarryItems(carryText);
        return blocks.Select(ParseNetBlock).Where(x => x.LengthM != null || x.WidthM != null || x.Specs.Count > 0).ToList();
    }

    private static List<string> SplitNetBlocks(string section, out string carryText)
    {
        carryText = string.Empty;
        var main = section;
        var carryMatch = Regex.Match(section, "携带[:：](.*)$", RegexOptions.Singleline);
        if (carryMatch.Success)
        {
            carryText = carryMatch.Groups[1].Value.Trim();
            main = section[..carryMatch.Index];
        }

        var matches = Regex.Matches(main, "(?:柔性网|硬质网)尺寸[:：]");
        if (matches.Count == 0) return string.IsNullOrWhiteSpace(main) ? [] : [main.Trim()];

        var list = new List<string>();
        for (var i = 0; i < matches.Count; i++)
        {
            var start = matches[i].Index;
            var end = i + 1 < matches.Count ? matches[i + 1].Index : main.Length;
            list.Add(main[start..end].Trim());
        }
        return list;
    }

    private static NetItem ParseNetBlock(string block)
    {
        var size = Regex.Match(block, @"(?:柔性网|硬质网)尺寸[:：]\s*长\s*(?<length>[\d.]+)\s*米，?\s*宽\s*(?<width>[\d.]+)\s*米，?\s*长度分\s*(?<segments>[\d.]+)\s*段");
        var specLine = Regex.Match(block, "规格[:：]\\s*([^\\n]+)");
        var mixed = ParseMixedWeaving(block);
        var width = ToDecimal(size.Groups["width"].Success ? size.Groups["width"].Value : null);

        return new NetItem
        {
            NetType = ClassifyNetType(block),
            ProductStandard = ClassifyProductStandard(block),
            LengthM = ToDecimal(size.Groups["length"].Success ? size.Groups["length"].Value : null),
            WidthM = width,
            Segments = ToInt(size.Groups["segments"].Success ? size.Groups["segments"].Value : null),
            Specs = specLine.Success ? ParseSpecs(specLine.Groups[1].Value) : [],
            IsMixedWeaving = specLine.Success && ParseSpecs(specLine.Groups[1].Value).Distinct().Count() > 1,
            MixedWeaving = mixed,
            RopeProcess = ParseRopeProcess(block),
            Validation = new NetValidation
            {
                MixedWidthSumM = mixed.Count == 0 ? null : mixed.Sum(x => x.WidthM ?? 0),
                MixedWidthMatchesNetWidth = width != null && mixed.Count > 0 && Math.Abs(mixed.Sum(x => x.WidthM ?? 0) - width.Value) < 0.001m
            }
        };
    }

    private static List<string> ParseSpecs(string raw)
    {
        return Regex.Matches(raw, "\\d+\\s*[*\\-]\\s*\\d+")
            .Select(m => NormalizeSpec(m.Value))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct()
            .Cast<string>()
            .ToList();
    }

    private static List<MixedWeavingItem> ParseMixedWeaving(string block)
    {
        var m = Regex.Match(block, "混编要求[:：](.*?)(?:整编钢丝绳要求|距（|距\\(|$)", RegexOptions.Singleline);
        if (!m.Success) return [];

        return Regex.Matches(m.Groups[1].Value, @"规格\s*(?<spec>[\d\s*\-]+).*?宽度\s*(?<width>[\d.]+)\s*米")
            .Select(x => new MixedWeavingItem
            {
                Spec = NormalizeSpec(x.Groups["spec"].Value),
                WidthM = ToDecimal(x.Groups["width"].Value)
            })
            .ToList();
    }

    private static RopeProcess ParseRopeProcess(string block)
    {
        var process = new RopeProcess
        {
            FirstRopeDistanceFromEdgeM = MatchDecimal(block, "距[（(]?\\s*网[边编]钢丝绳\\s*[）)]?位置[（(]?\\s*([\\d.]+)\\s*[）)]?米编第一根"),
            ExposedEachEndM = MatchDecimal(block, "钢丝绳两端外露\\s*([\\d.]+)\\s*米")
        };

        foreach (Match m in Regex.Matches(block,
                     @"(?:间距\s*)?(?<spacing>[\d.]+)\s*(?:米|m)?\s*共编\s*(?<count>\d+)\s*根?\s*(?:整编)?(?:一般)?钢丝绳\s*(?<spec>[Φφ]\s*[\d.]+)?",
                     RegexOptions.IgnoreCase))
        {
            process.Items.Add(new RopeProcessItem
            {
                SpacingM = ToDecimal(m.Groups["spacing"].Value),
                Count = ToInt(m.Groups["count"].Value),
                RopeType = "整编一般钢丝绳",
                Spec = NormalizeSpec(m.Groups["spec"].Success ? m.Groups["spec"].Value : null)
            });
        }

        return process;
    }

    private static List<CarryItem> ParseCarryItems(string carryText)
    {
        if (string.IsNullOrWhiteSpace(carryText)) return [];
        var list = new List<CarryItem>();
        foreach (var part in Regex.Split(carryText, "[；;]\\s*"))
        {
            var text = part.Trim();
            if (text.Length == 0) continue;

            var net = Regex.Match(text, @"(?<name>柔性网|硬质网)规格\s*(?<spec>[\d\s*\-]+)\s*长\s*(?<length>[\d.]+)\s*米\s*，?\s*宽\s*(?<width>[\d.]+)\s*m?\s*，?\s*片数\s*(?<pieces>\d+)");
            if (net.Success)
            {
                list.Add(new CarryItem
                {
                    Name = Clean(net.Groups["name"].Value),
                    Spec = NormalizeSpec(net.Groups["spec"].Value),
                    LengthM = ToDecimal(net.Groups["length"].Value),
                    WidthM = ToDecimal(net.Groups["width"].Value),
                    Pieces = ToInt(net.Groups["pieces"].Value),
                    ShipDirectly = true
                });
                continue;
            }

            var rope = Regex.Match(text, @"一般钢丝绳\s*(?<spec>[Φφ]\s*[\d.]+).*?长度\s*(?<length>[\d.]+)\s*米(?:单价\s*(?<price>[\d.]+)\s*元)?");
            if (rope.Success)
            {
                var detail = Regex.Match(text, "具体为[:：]\\s*(.*)");
                list.Add(new CarryItem
                {
                    Name = "一般钢丝绳",
                    Spec = NormalizeSpec(rope.Groups["spec"].Value),
                    TotalLengthM = ToDecimal(rope.Groups["length"].Value),
                    UnitPrice = ToDecimal(rope.Groups["price"].Success ? rope.Groups["price"].Value : null),
                    Detail = detail.Success ? detail.Groups[1].Value.Trim() : null,
                    ShipDirectly = true
                });
            }
        }
        return list;
    }

    private static List<UndergroundTransportSize> ParseUndergroundTransport(string section)
    {
        return Regex.Matches(section, @"长\s*(?<length>[\d.]+)\s*m.*?(?:金额|单价)\s*(?<amount>[\d.]+)", RegexOptions.Singleline)
            .Select(m => new UndergroundTransportSize
            {
                LengthM = ToDecimal(m.Groups["length"].Value),
                AmountOrUnitPrice = ToDecimal(m.Groups["amount"].Value)
            })
            .ToList();
    }

    private static List<AccessoryItem> ParseAccessories(string section)
    {
        var skip = new HashSet<string> { "产品及配件", "产品", "型号", "数量", "单价" };
        var cells = section.Split('\n')
            .Select(x => x.Trim())
            .Where(x => x.Length > 0 && !skip.Contains(x))
            .ToList();

        var list = new List<AccessoryItem>();
        for (var i = 0; i + 2 < cells.Count;)
        {
            var name = cells[i];
            var model = cells[i + 1];
            var qty = cells[i + 2];
            if (IsNumber(qty))
            {
                decimal? unitPrice = null;
                var step = 3;
                if (i + 4 < cells.Count && IsNumber(cells[i + 3]) && !IsNumber(cells[i + 4]))
                {
                    unitPrice = ToDecimal(cells[i + 3]);
                    step = 4;
                }

                list.Add(new AccessoryItem
                {
                    Name = name,
                    Model = CleanModel(model),
                    Quantity = ToDecimal(qty),
                    UnitPrice = unitPrice
                });
                i += step;
            }
            else
            {
                i++;
            }
        }
        return list;
    }

    private static SalesInfo ParseSalesFooter(string text)
    {
        var m = Regex.Match(text, "销售部门[:：]\\s*(.*?)\\s*销售经理[:：]\\s*(.*?)\\s*(?:审\\s*核|日期|日\\s*期|$)", RegexOptions.Singleline);
        return new SalesInfo
        {
            SalesDepartment = m.Success ? m.Groups[1].Value.Trim() : null,
            SalesManager = m.Success ? m.Groups[2].Value.Trim() : null,
        };
    }

    public static ValidationResult Validate(ParsedOrder order)
    {
        var result = new ValidationResult();
        foreach (var (net, index) in order.Nets.Select((n, i) => (n, i + 1)))
        {
            var mixedWidthSum = net.MixedWeaving.Count == 0 ? (decimal?)null : net.MixedWeaving.Sum(x => x.WidthM ?? 0);
            net.Validation.MixedWidthSumM = mixedWidthSum;
            net.Validation.MixedWidthMatchesNetWidth = net.WidthM != null && mixedWidthSum != null && Math.Abs(mixedWidthSum.Value - net.WidthM.Value) < 0.001m;
            if (net.Specs.Count > 1 && net.MixedWeaving.Count == 0)
                result.Warnings.Add($"第 {index} 个网有多个规格，但未解析到混编要求。");
            if (net.Specs.Count <= 1 && net.MixedWeaving.Count > 0)
                result.Warnings.Add($"第 {index} 个网只有一个规格，但出现了混编要求，请人工确认。");
            if (net.MixedWeaving.Count > 0 && net.Validation?.MixedWidthMatchesNetWidth == false)
                result.Warnings.Add($"第 {index} 个网混编宽度合计与网宽不一致。");
        }
        return result;
    }

    private static string? ClassifyNetType(string block)
    {
        if (block.Contains("柔性网") || block.Contains("聚酯") || block.Contains("防护网")) return "柔性网";
        if (block.Contains("硬质网") || block.Contains("护帮网") || block.Contains("川藏网") || block.Contains("非阻燃硬质网")) return "硬质网";
        return null;
    }

    private static string? ClassifyProductStandard(string block)
    {
        var candidates = new[] { "煤矿井下用聚酯纤维防护网", "煤矿井下用聚酯防护网", "硬质网用护帮网", "硬质川藏网", "非阻燃硬质网" };
        return candidates.FirstOrDefault(block.Contains);
    }

    private static string? ParseChineseDate(string text)
    {
        var m = Regex.Match(text, "发货时间[:：]?\\s*(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日");
        return m.Success ? $"{int.Parse(m.Groups[1].Value):0000}-{int.Parse(m.Groups[2].Value):00}-{int.Parse(m.Groups[3].Value):00}" : null;
    }

    private static string? MatchValue(string text, string pattern)
    {
        var m = Regex.Match(text, pattern, RegexOptions.Singleline);
        return m.Success ? Clean(m.Groups[1].Value) : null;
    }

    private static decimal? MatchDecimal(string text, string pattern)
    {
        var m = Regex.Match(text, pattern, RegexOptions.Singleline);
        return m.Success ? ToDecimal(m.Groups[1].Value) : null;
    }

    private static decimal? ToDecimal(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        return decimal.TryParse(value.Replace(",", "").Trim(), out var d) ? d : null;
    }

    private static int? ToInt(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        return decimal.TryParse(value.Trim(), out var d) ? (int)d : null;
    }

    private static string? NormalizeSpec(string? spec)
    {
        if (string.IsNullOrWhiteSpace(spec)) return null;
        spec = Regex.Replace(spec, "\\s+", "").Replace('-', '*').Replace('Φ', 'φ');
        return spec;
    }

    private static string Clean(string value) => Regex.Replace(value, "\\s+", " ").Trim();
    private static string CleanModel(string value) => Clean(value.Replace("__", "")).Trim('_');
    private static bool IsNumber(string value) => Regex.IsMatch(value.Trim(), "^[\\d.]+$");
}

public static class YxAiClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true
    };

    public static async Task<AiReviewResult> FallbackAndReviewAsync(string baseUrl, string apiKey, string model, string apiFormat, string rawText, ParsedOrder parsed)
    {
        var review = new AiReviewResult { Model = model, ApiFormat = apiFormat };
        using var http = CreateHttpClient(apiKey);
        var endpoint = baseUrl.TrimEnd('/');

        foreach (var task in BuildFallbackTasks(rawText, parsed))
        {
            var aiText = await ChatAsync(http, endpoint, model, apiFormat, BuildFallbackSystemPrompt(), task.UserPrompt);
            review.SectionResults.Add(new AiSectionResult
            {
                Section = task.Section,
                Target = task.Target,
                Raw = aiText
            });

            var fallback = TryParseJson<AiFallbackResponse>(aiText, review);
            if (fallback == null) continue;

            review.Warnings.AddRange(fallback.Warnings ?? []);
            review.Suggestions.AddRange(fallback.Suggestions ?? []);
            ApplyFallback(parsed, fallback, review, task.TargetNetIndex);
        }

        var reviewText = await ChatAsync(http, endpoint, model, apiFormat, BuildReviewSystemPrompt(), BuildReviewUserPrompt(rawText, parsed));
        review.Raw = reviewText;
        var finalReview = TryParseJson<AiReviewJson>(reviewText, review);
        if (finalReview != null)
        {
            review.Warnings.AddRange(finalReview.Warnings ?? []);
            review.Suggestions.AddRange(finalReview.Suggestions ?? []);
        }

        review.Warnings = review.Warnings.Distinct().ToList();
        review.Suggestions = review.Suggestions.Distinct().ToList();
        return review;
    }

    private static HttpClient CreateHttpClient(string apiKey)
    {
        var http = new HttpClient { Timeout = TimeSpan.FromSeconds(120) };
        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        return http;
    }

    private static async Task<string> ChatAsync(HttpClient http, string baseEndpoint, string model, string apiFormat, string systemPrompt, string userPrompt)
    {
        return apiFormat.Equals("claude_messages", StringComparison.OrdinalIgnoreCase)
            ? await ClaudeMessagesAsync(http, baseEndpoint, model, systemPrompt, userPrompt)
            : await OpenAiChatAsync(http, baseEndpoint, model, systemPrompt, userPrompt);
    }

    private static async Task<string> OpenAiChatAsync(HttpClient http, string baseEndpoint, string model, string systemPrompt, string userPrompt)
    {
        var endpoint = baseEndpoint.TrimEnd('/') + "/chat/completions";
        var body = new
        {
            model,
            temperature = 0,
            messages = new object[]
            {
                new { role = "system", content = systemPrompt },
                new { role = "user", content = userPrompt }
            }
        };

        using var response = await http.PostAsync(endpoint, new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"));
        var content = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode) return ErrorJson(response, content);

        try
        {
            using var doc = JsonDocument.Parse(content);
            return doc.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? "{}";
        }
        catch
        {
            return JsonSerializer.Serialize(new { warnings = new[] { "AI 返回格式不是预期的 OpenAI Chat Completions JSON。" }, raw_response = content }, JsonOptions);
        }
    }

    private static async Task<string> ClaudeMessagesAsync(HttpClient http, string baseEndpoint, string model, string systemPrompt, string userPrompt)
    {
        var endpoint = baseEndpoint.TrimEnd('/') + "/messages";
        var body = new
        {
            model,
            max_tokens = 4096,
            temperature = 0,
            system = systemPrompt,
            messages = new object[]
            {
                new { role = "user", content = userPrompt }
            }
        };

        using var response = await http.PostAsync(endpoint, new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"));
        var content = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode) return ErrorJson(response, content);

        try
        {
            using var doc = JsonDocument.Parse(content);
            var contentElement = doc.RootElement.GetProperty("content");
            if (contentElement.ValueKind == JsonValueKind.Array)
            {
                var sb = new StringBuilder();
                foreach (var item in contentElement.EnumerateArray())
                {
                    if (item.TryGetProperty("text", out var text)) sb.Append(text.GetString());
                }
                return sb.ToString();
            }
            return contentElement.ToString();
        }
        catch
        {
            return JsonSerializer.Serialize(new { warnings = new[] { "AI 返回格式不是预期的 Claude Messages JSON。" }, raw_response = content }, JsonOptions);
        }
    }

    private static string ErrorJson(HttpResponseMessage response, string content)
    {
        return JsonSerializer.Serialize(new
        {
            warnings = new[] { $"AI 接口请求失败：HTTP {(int)response.StatusCode}" },
            raw_error = content
        }, JsonOptions);
    }

    private static List<AiFallbackTask> BuildFallbackTasks(string rawText, ParsedOrder parsed)
    {
        var tasks = new List<AiFallbackTask>();
        var sections = ExtractSectionsForAi(rawText);
        var section1 = sections.GetValueOrDefault("1") ?? string.Empty;
        var section3 = sections.GetValueOrDefault("3") ?? string.Empty;
        var section6 = sections.GetValueOrDefault("6") ?? string.Empty;
        var accessory = sections.Values.FirstOrDefault(x => x.Contains("产品及配件", StringComparison.Ordinal)) ?? sections.GetValueOrDefault("11") ?? sections.GetValueOrDefault("13") ?? string.Empty;

        if (parsed.BasicInfo.Customer == null || parsed.BasicInfo.SalesType == null || parsed.BasicInfo.ContractTotalAmount == null || parsed.BasicInfo.DeliveryDate == null || parsed.BasicInfo.PriceTerms.Count == 0)
        {
            tasks.Add(new AiFallbackTask("1", "basic_info", null, BuildBasicInfoPrompt(section1, parsed.BasicInfo)));
        }

        if (parsed.Nets.Count == 0 && !string.IsNullOrWhiteSpace(section3))
        {
            tasks.Add(new AiFallbackTask("3", "nets", null, BuildNetsPrompt(section3, parsed.Nets)));
        }
        else
        {
            for (var i = 0; i < parsed.Nets.Count; i++)
            {
                var net = parsed.Nets[i];
                var lowConfidence = net.LengthM == null || net.WidthM == null || net.Segments == null || net.Specs.Count == 0 ||
                                    (net.Specs.Count > 1 && net.MixedWeaving.Count == 0) ||
                                    net.Validation.MixedWidthMatchesNetWidth == false;
                if (lowConfidence)
                {
                    tasks.Add(new AiFallbackTask("3", "net", i, BuildSingleNetPrompt(section3, i, net)));
                }
            }
        }

        if (parsed.UndergroundTransportSizes.Count == 0 && !string.IsNullOrWhiteSpace(section6))
        {
            tasks.Add(new AiFallbackTask("6", "underground_transport_sizes", null, BuildTransportPrompt(section6)));
        }

        if (parsed.Accessories.Count == 0 && !string.IsNullOrWhiteSpace(accessory))
        {
            tasks.Add(new AiFallbackTask("accessories", "accessories", null, BuildAccessoriesPrompt(accessory)));
        }

        return tasks;
    }

    private static void ApplyFallback(ParsedOrder parsed, AiFallbackResponse fallback, AiReviewResult review, int? targetNetIndex)
    {
        if (fallback.BasicInfo != null)
        {
            MergeBasicInfo(parsed.BasicInfo, fallback.BasicInfo, review);
        }

        if (fallback.Nets is { Count: > 0 })
        {
            if (targetNetIndex != null && targetNetIndex.Value >= 0 && targetNetIndex.Value < parsed.Nets.Count)
            {
                MergeNet(parsed.Nets[targetNetIndex.Value], fallback.Nets[0], review, targetNetIndex.Value);
            }
            else if (parsed.Nets.Count == 0)
            {
                parsed.Nets.AddRange(fallback.Nets.Select(x =>
                {
                    x.ApplySource("ai_fallback");
                    return x;
                }));
                review.AppliedPatches.Add("AI 补全 nets：规则未识别到网块，已采用 AI 返回的网块。请人工重点复核。");
            }
        }

        if (fallback.UndergroundTransportSizes is { Count: > 0 } && parsed.UndergroundTransportSizes.Count == 0)
        {
            parsed.UndergroundTransportSizes = fallback.UndergroundTransportSizes;
            review.AppliedPatches.Add("AI 补全 underground_transport_sizes。请核对原文依据。");
        }

        if (fallback.Accessories is { Count: > 0 } && parsed.Accessories.Count == 0)
        {
            parsed.Accessories = fallback.Accessories;
            review.AppliedPatches.Add("AI 补全 accessories。请核对表格列是否错位。");
        }
    }

    private static void MergeBasicInfo(BasicInfo target, BasicInfo patch, AiReviewResult review)
    {
        if (target.Customer == null && patch.Customer != null) { target.Customer = patch.Customer; review.AppliedPatches.Add("AI 补全 basic_info.customer"); }
        if (target.SalesType == null && patch.SalesType != null) { target.SalesType = patch.SalesType; review.AppliedPatches.Add("AI 补全 basic_info.sales_type"); }
        if (target.ContractTotalAmount == null && patch.ContractTotalAmount != null) { target.ContractTotalAmount = patch.ContractTotalAmount; review.AppliedPatches.Add("AI 补全 basic_info.contract_total_amount"); }
        if (target.ExchangeRate == null && patch.ExchangeRate != null) { target.ExchangeRate = patch.ExchangeRate; review.AppliedPatches.Add("AI 补全 basic_info.exchange_rate"); }
        if (target.DeliveryDate == null && patch.DeliveryDate != null) { target.DeliveryDate = patch.DeliveryDate; review.AppliedPatches.Add("AI 补全 basic_info.delivery_date"); }
        if (target.PriceTerms.Count == 0 && patch.PriceTerms.Count > 0) { target.PriceTerms = patch.PriceTerms; review.AppliedPatches.Add("AI 补全 basic_info.price_terms"); }
    }

    private static void MergeNet(NetItem target, NetItem patch, AiReviewResult review, int index)
    {
        var prefix = $"nets[{index}]";
        if (target.NetType == null && patch.NetType != null) { target.NetType = patch.NetType; review.AppliedPatches.Add($"AI 补全 {prefix}.net_type"); }
        if (target.ProductStandard == null && patch.ProductStandard != null) { target.ProductStandard = patch.ProductStandard; review.AppliedPatches.Add($"AI 补全 {prefix}.product_standard"); }
        if (target.LengthM == null && patch.LengthM != null) { target.LengthM = patch.LengthM; review.AppliedPatches.Add($"AI 补全 {prefix}.length_m"); }
        if (target.WidthM == null && patch.WidthM != null) { target.WidthM = patch.WidthM; review.AppliedPatches.Add($"AI 补全 {prefix}.width_m"); }
        if (target.Segments == null && patch.Segments != null) { target.Segments = patch.Segments; review.AppliedPatches.Add($"AI 补全 {prefix}.segments"); }
        if (target.Specs.Count == 0 && patch.Specs.Count > 0) { target.Specs = patch.Specs; review.AppliedPatches.Add($"AI 补全 {prefix}.specs"); }
        if (target.MixedWeaving.Count == 0 && patch.MixedWeaving.Count > 0) { target.MixedWeaving = patch.MixedWeaving; review.AppliedPatches.Add($"AI 补全 {prefix}.mixed_weaving"); }
        if (target.RopeProcess.FirstRopeDistanceFromEdgeM == null && patch.RopeProcess.FirstRopeDistanceFromEdgeM != null) { target.RopeProcess.FirstRopeDistanceFromEdgeM = patch.RopeProcess.FirstRopeDistanceFromEdgeM; review.AppliedPatches.Add($"AI 补全 {prefix}.rope_process.first_rope_distance_from_edge_m"); }
        if (target.RopeProcess.Items.Count == 0 && patch.RopeProcess.Items.Count > 0) { target.RopeProcess.Items = patch.RopeProcess.Items; review.AppliedPatches.Add($"AI 补全 {prefix}.rope_process.items"); }
        if (target.RopeProcess.ExposedEachEndM == null && patch.RopeProcess.ExposedEachEndM != null) { target.RopeProcess.ExposedEachEndM = patch.RopeProcess.ExposedEachEndM; review.AppliedPatches.Add($"AI 补全 {prefix}.rope_process.exposed_each_end_m"); }
        target.IsMixedWeaving = target.Specs.Count > 1 || target.MixedWeaving.Count > 0;
        target.ApplySource("rule+ai_fallback");
    }

    private static T? TryParseJson<T>(string text, AiReviewResult review)
    {
        try
        {
            var json = ExtractJsonObject(text);
            return JsonSerializer.Deserialize<T>(json, JsonOptions);
        }
        catch (Exception ex)
        {
            review.Warnings.Add("AI 返回内容不是可解析 JSON：" + ex.Message);
            return default;
        }
    }

    private static string ExtractJsonObject(string text)
    {
        var trimmed = text.Trim();
        if (trimmed.StartsWith("```"))
        {
            trimmed = Regex.Replace(trimmed, "^```(?:json)?", "", RegexOptions.IgnoreCase).Trim();
            trimmed = Regex.Replace(trimmed, "```$", "").Trim();
        }
        var start = trimmed.IndexOf('{');
        var end = trimmed.LastIndexOf('}');
        if (start >= 0 && end > start) return trimmed[start..(end + 1)];
        return trimmed;
    }

    private static Dictionary<string, string> ExtractSectionsForAi(string rawText)
    {
        var text = DocumentTextExtractor.Normalize(rawText);
        var result = new Dictionary<string, List<string>>();
        string? current = null;
        var lastNo = 0;
        foreach (var raw in text.Split('\n'))
        {
            var line = raw.Trim();
            if (Regex.IsMatch(line, "^\\d{1,2}$") && int.TryParse(line, out var no) && no >= 1 && no <= 14 && no > lastNo)
            {
                current = no.ToString();
                lastNo = no;
                result.TryAdd(current, []);
                continue;
            }
            if (current != null && line.Length > 0) result[current].Add(line);
        }
        return result.ToDictionary(kv => kv.Key, kv => string.Join('\n', kv.Value));
    }

    private static string BuildFallbackSystemPrompt() => """
        你是矿用网订单解析助手。你的任务是对规则解析失败或低置信度的局部章节做补全。
        严格要求：
        1. 只根据用户提供的原文片段抽取，不要猜测，不要套常识补值。
        2. 只返回 JSON object，不要 Markdown，不要解释文字。
        3. 每个补充值必须能在原文中找到依据；不确定就返回 null 或空数组，并写入 warnings。
        4. 数值统一用米、元、根、片等基础数值，不带单位字符串。
        5. 规格统一为 600*600、300*300、φ21.5 这类格式。
        6. 网编/网边钢丝绳不要作为目标字段，只解析整编一般钢丝绳工艺。
        """;

    private static string BuildReviewSystemPrompt() => """
        你是矿用网订单解析审查助手。请审查最终 JSON 与原文是否矛盾、是否漏掉关键业务字段。
        只返回 JSON object，格式：{"warnings":[],"suggestions":[]}。
        不要编造原文不存在的信息，不要输出 Markdown。
        """;

    private static string BuildBasicInfoPrompt(string section, BasicInfo current) =>
        "请从“序号1”原文中补全 basic_info。\n" +
        "返回格式：{\"basic_info\":{\"customer\":null,\"sales_type\":null,\"contract_total_amount\":null,\"exchange_rate\":null,\"price_terms\":[{\"name\":null,\"spec\":null,\"unit_price\":null,\"amount\":null,\"pricing_method\":null}],\"delivery_date\":null},\"warnings\":[],\"suggestions\":[]}\n\n" +
        "当前规则结果：\n" + JsonSerializer.Serialize(current, JsonOptions) + "\n\n" +
        "原文：\n" + Truncate(section, 6000);

    private static string BuildNetsPrompt(string section, List<NetItem> current) =>
        "请从“序号3”原文中解析所有 nets。不要解析携带成品；携带成品不是加工网。\n" +
        "返回格式：{\"nets\":[{\"net_type\":null,\"product_standard\":null,\"length_m\":null,\"width_m\":null,\"segments\":null,\"specs\":[],\"is_mixed_weaving\":false,\"mixed_weaving\":[{\"spec\":null,\"width_m\":null,\"evidence\":null}],\"rope_process\":{\"first_rope_distance_from_edge_m\":null,\"items\":[{\"spacing_m\":null,\"count\":null,\"rope_type\":\"整编一般钢丝绳\",\"spec\":null,\"evidence\":null}],\"exposed_each_end_m\":null}}],\"warnings\":[],\"suggestions\":[]}\n\n" +
        "当前规则结果：\n" + JsonSerializer.Serialize(current, JsonOptions) + "\n\n" +
        "原文：\n" + Truncate(section, 10000);

    private static string BuildSingleNetPrompt(string section, int index, NetItem current) =>
        $"请只补全第 {index + 1} 个网的缺失字段。不要覆盖规则已正确解析的字段。\n" +
        "返回格式：{\"nets\":[{\"net_type\":null,\"product_standard\":null,\"length_m\":null,\"width_m\":null,\"segments\":null,\"specs\":[],\"is_mixed_weaving\":false,\"mixed_weaving\":[{\"spec\":null,\"width_m\":null,\"evidence\":null}],\"rope_process\":{\"first_rope_distance_from_edge_m\":null,\"items\":[{\"spacing_m\":null,\"count\":null,\"rope_type\":\"整编一般钢丝绳\",\"spec\":null,\"evidence\":null}],\"exposed_each_end_m\":null}}],\"warnings\":[],\"suggestions\":[]}\n\n" +
        $"当前第 {index + 1} 个网规则结果：\n" + JsonSerializer.Serialize(current, JsonOptions) + "\n\n" +
        "序号3原文：\n" + Truncate(section, 10000);

    private static string BuildTransportPrompt(string section) =>
        "请从序号6原文中只抽取井下运输尺寸的“长”和“金额/单价”。\n" +
        "返回格式：{\"underground_transport_sizes\":[{\"length_m\":null,\"amount_or_unit_price\":null}],\"warnings\":[],\"suggestions\":[]}\n\n" +
        "原文：\n" + Truncate(section, 6000);

    private static string BuildAccessoriesPrompt(string section) =>
        "请从产品及配件表原文中抽取配件列表，字段为 name、model、quantity、unit_price。没有单价则 unit_price 为 null。\n" +
        "返回格式：{\"accessories\":[{\"name\":null,\"model\":null,\"quantity\":null,\"unit_price\":null}],\"warnings\":[],\"suggestions\":[]}\n\n" +
        "原文：\n" + Truncate(section, 10000);

    private static string BuildReviewUserPrompt(string rawText, ParsedOrder parsed) => "原文：\n" + Truncate(rawText, 12000) + "\n\n最终解析 JSON：\n" + JsonSerializer.Serialize(parsed, JsonOptions);

    private static string Truncate(string value, int maxChars) => value.Length <= maxChars ? value : value[..maxChars];
}

public sealed record AiFallbackTask(string Section, string Target, int? TargetNetIndex, string UserPrompt);

public sealed class AiFallbackResponse
{
    [JsonPropertyName("basic_info")] public BasicInfo? BasicInfo { get; set; }
    [JsonPropertyName("nets")] public List<NetItem>? Nets { get; set; }
    [JsonPropertyName("underground_transport_sizes")] public List<UndergroundTransportSize>? UndergroundTransportSizes { get; set; }
    [JsonPropertyName("accessories")] public List<AccessoryItem>? Accessories { get; set; }
    [JsonPropertyName("warnings")] public List<string>? Warnings { get; set; }
    [JsonPropertyName("suggestions")] public List<string>? Suggestions { get; set; }
}

public sealed class AiReviewJson
{
    [JsonPropertyName("warnings")] public List<string>? Warnings { get; set; }
    [JsonPropertyName("suggestions")] public List<string>? Suggestions { get; set; }
}

public sealed class ParsedOrder
{
    [JsonPropertyName("basic_info")] public BasicInfo BasicInfo { get; set; } = new();
    [JsonPropertyName("nets")] public List<NetItem> Nets { get; set; } = [];
    [JsonPropertyName("carry_items")] public List<CarryItem> CarryItems { get; set; } = [];
    [JsonPropertyName("underground_transport_sizes")] public List<UndergroundTransportSize> UndergroundTransportSizes { get; set; } = [];
    [JsonPropertyName("accessories")] public List<AccessoryItem> Accessories { get; set; } = [];
    [JsonPropertyName("sales")] public SalesInfo Sales { get; set; } = new();
    [JsonPropertyName("validation")] public ValidationResult Validation { get; set; } = new();
    [JsonPropertyName("ai_review")] public AiReviewResult? AiReview { get; set; }
    [JsonPropertyName("raw_section_keys")] public List<string> RawSectionKeys { get; set; } = [];
}

public sealed class BasicInfo
{
    [JsonPropertyName("customer")] public string? Customer { get; set; }
    [JsonPropertyName("sales_type")] public string? SalesType { get; set; }
    [JsonPropertyName("contract_total_amount")] public decimal? ContractTotalAmount { get; set; }
    [JsonPropertyName("exchange_rate")] public decimal? ExchangeRate { get; set; }
    [JsonPropertyName("price_terms")] public List<PriceTerm> PriceTerms { get; set; } = [];
    [JsonPropertyName("delivery_date")] public string? DeliveryDate { get; set; }
}

public sealed class PriceTerm
{
    [JsonPropertyName("name")] public string? Name { get; set; }
    [JsonPropertyName("spec")] public string? Spec { get; set; }
    [JsonPropertyName("unit_price")] public decimal? UnitPrice { get; set; }
    [JsonPropertyName("amount")] public decimal? Amount { get; set; }
    [JsonPropertyName("pricing_method")] public string? PricingMethod { get; set; }
}

public sealed class NetItem
{
    [JsonPropertyName("net_type")] public string? NetType { get; set; }
    [JsonPropertyName("product_standard")] public string? ProductStandard { get; set; }
    [JsonPropertyName("length_m")] public decimal? LengthM { get; set; }
    [JsonPropertyName("width_m")] public decimal? WidthM { get; set; }
    [JsonPropertyName("segments")] public int? Segments { get; set; }
    [JsonPropertyName("specs")] public List<string> Specs { get; set; } = [];
    [JsonPropertyName("is_mixed_weaving")] public bool IsMixedWeaving { get; set; }
    [JsonPropertyName("mixed_weaving")] public List<MixedWeavingItem> MixedWeaving { get; set; } = [];
    [JsonPropertyName("rope_process")] public RopeProcess RopeProcess { get; set; } = new();
    [JsonPropertyName("validation")] public NetValidation Validation { get; set; } = new();
    [JsonPropertyName("source")] public string? Source { get; set; }

    public void ApplySource(string source)
    {
        Source = source;
        foreach (var item in MixedWeaving) item.Source ??= source;
        RopeProcess.Source ??= source;
        foreach (var item in RopeProcess.Items) item.Source ??= source;
    }
}

public sealed class MixedWeavingItem
{
    [JsonPropertyName("spec")] public string? Spec { get; set; }
    [JsonPropertyName("width_m")] public decimal? WidthM { get; set; }
    [JsonPropertyName("source")] public string? Source { get; set; }
    [JsonPropertyName("evidence")] public string? Evidence { get; set; }
}

public sealed class RopeProcess
{
    [JsonPropertyName("first_rope_distance_from_edge_m")] public decimal? FirstRopeDistanceFromEdgeM { get; set; }
    [JsonPropertyName("items")] public List<RopeProcessItem> Items { get; set; } = [];
    [JsonPropertyName("exposed_each_end_m")] public decimal? ExposedEachEndM { get; set; }
    [JsonPropertyName("source")] public string? Source { get; set; }
}

public sealed class RopeProcessItem
{
    [JsonPropertyName("spacing_m")] public decimal? SpacingM { get; set; }
    [JsonPropertyName("count")] public int? Count { get; set; }
    [JsonPropertyName("rope_type")] public string? RopeType { get; set; }
    [JsonPropertyName("spec")] public string? Spec { get; set; }
    [JsonPropertyName("source")] public string? Source { get; set; }
    [JsonPropertyName("evidence")] public string? Evidence { get; set; }
}

public sealed class CarryItem
{
    [JsonPropertyName("name")] public string? Name { get; set; }
    [JsonPropertyName("spec")] public string? Spec { get; set; }
    [JsonPropertyName("length_m")] public decimal? LengthM { get; set; }
    [JsonPropertyName("width_m")] public decimal? WidthM { get; set; }
    [JsonPropertyName("pieces")] public int? Pieces { get; set; }
    [JsonPropertyName("total_length_m")] public decimal? TotalLengthM { get; set; }
    [JsonPropertyName("unit_price")] public decimal? UnitPrice { get; set; }
    [JsonPropertyName("detail")] public string? Detail { get; set; }
    [JsonPropertyName("ship_directly")] public bool ShipDirectly { get; set; }
}

public sealed class UndergroundTransportSize
{
    [JsonPropertyName("length_m")] public decimal? LengthM { get; set; }
    [JsonPropertyName("amount_or_unit_price")] public decimal? AmountOrUnitPrice { get; set; }
}

public sealed class AccessoryItem
{
    [JsonPropertyName("name")] public string? Name { get; set; }
    [JsonPropertyName("model")] public string? Model { get; set; }
    [JsonPropertyName("quantity")] public decimal? Quantity { get; set; }
    [JsonPropertyName("unit_price")] public decimal? UnitPrice { get; set; }
}

public sealed class SalesInfo
{
    [JsonPropertyName("sales_department")] public string? SalesDepartment { get; set; }
    [JsonPropertyName("sales_manager")] public string? SalesManager { get; set; }
}

public sealed class ValidationResult
{
    [JsonPropertyName("warnings")] public List<string> Warnings { get; set; } = [];
}

public sealed class NetValidation
{
    [JsonPropertyName("mixed_width_sum_m")] public decimal? MixedWidthSumM { get; set; }
    [JsonPropertyName("mixed_width_matches_net_width")] public bool? MixedWidthMatchesNetWidth { get; set; }
}

public sealed class AiReviewResult
{
    [JsonPropertyName("model")] public string? Model { get; set; }
    [JsonPropertyName("api_format")] public string? ApiFormat { get; set; }
    [JsonPropertyName("warnings")] public List<string> Warnings { get; set; } = [];
    [JsonPropertyName("suggestions")] public List<string> Suggestions { get; set; } = [];
    [JsonPropertyName("applied_patches")] public List<string> AppliedPatches { get; set; } = [];
    [JsonPropertyName("section_results")] public List<AiSectionResult> SectionResults { get; set; } = [];
    [JsonPropertyName("raw")] public string? Raw { get; set; }
}

public sealed class AiSectionResult
{
    [JsonPropertyName("section")] public string? Section { get; set; }
    [JsonPropertyName("target")] public string? Target { get; set; }
    [JsonPropertyName("raw")] public string? Raw { get; set; }
}

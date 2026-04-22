using System.Text.Json;

namespace AIWorkAssistant.Services.Agent;

/// <summary>
/// 用 AI 解析 Word 文档内容为结构化 JSON
/// </summary>
public class OrderParserService
{
    private readonly ChatService _chatService;

    public OrderParserService(ChatService chatService)
    {
        _chatService = chatService;
    }

    public async Task<JsonDocument> ParseOrderAsync(string documentText)
    {
        var prompt = """
            你是一个订单数据解析专家。请将以下"格瑞特矿用聚酯纤维柔性网订货通知单"的内容解析为结构化的 JSON。

            要求：
            1. 提取以下关键字段：
               - customerName: 使用单位名称
               - contractAmount: 合同金额（数字）
               - workFaceName: 工作面名称
               - deliveryDate: 到货时间（如果有）
               - workFaceAngle: 工作面倾角（数字）
               - miningHeight: 工作面采高（数字）
               - stopMiningHeight: 停采时采高（数字）
               - netSpec: 柔性网规格型号
               - netLength: 柔性网长度（数字，米）
               - netWidth: 柔性网宽度（数字，米）
               - netSegments: 长度分几段（数字）
               - wireRopeSpec: 整编钢丝绳规格
               - wireRopeLength: 整编钢丝绳长度（数字，米）
               - wireRopeAmount: 整编钢丝绳金额（数字）
               - transportDirection: 运输展开方向（左进/右进）
               - transportMethod: 运输方式（立井/斜井/平硐）
               - maxTransportSize: 井下运输最大尺寸（对象：length, width, height）
               - transportFrameAmount: 运输框架金额（数字）
               - supportCount: 支架个数（数字）
               - remark: 其他备注信息
            2. products: 产品及配件清单，数组格式，每项包含：
               - name: 产品名称
               - spec: 型号/规格
               - quantity: 数量（数字）
               - unit: 单位
               - unitPrice: 单价（数字，如果有）
               - amount: 金额（数字，如果有）
            3. netProducts: 柔性网产品清单（从第3项中提取），每项包含：
               - spec: 规格型号（如 H PET 700-700MS）
               - width: 宽度（数字，米）
               - length: 长度（数字，米）
               - segments: 段数（数字）
               - amount: 金额（数字，如果有）
            4. 数值类型用数字，空白未填写的用 null
            5. 只返回 JSON，不要返回其他内容

            文档内容：
            """ + documentText;

        var messages = new List<(string role, string content)>
        {
            ("user", prompt)
        };

        var response = await _chatService.SendMessageAsync("", messages);
        var jsonText = ExtractJson(response);
        return JsonDocument.Parse(jsonText);
    }

    private static string ExtractJson(string text)
    {
        var jsonStart = text.IndexOf("```json", StringComparison.OrdinalIgnoreCase);
        if (jsonStart >= 0)
        {
            jsonStart = text.IndexOf('\n', jsonStart) + 1;
            var jsonEnd = text.IndexOf("```", jsonStart, StringComparison.Ordinal);
            if (jsonEnd > jsonStart)
                text = text[jsonStart..jsonEnd].Trim();
        }
        else
        {
            var first = text.IndexOf('{');
            var last = text.LastIndexOf('}');
            if (first >= 0 && last > first)
                text = text[first..(last + 1)];
        }

        return text;
    }
}

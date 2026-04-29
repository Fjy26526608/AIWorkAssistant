using System.Globalization;
using System.Text;
using AIWorkAssistant.Models.HkOrder;

namespace AIWorkAssistant.Services.OrderParsing;

public static class MiningOrderUploadMapper
{
    public static OrderData ToUploadOrder(ParsedOrder parsed)
    {
        var deliveryDate = parsed.BasicInfo.DeliveryDate ?? DateTime.Today.AddDays(15).ToString("yyyy-MM-dd");
        var priceTerms = parsed.BasicInfo.PriceTerms;

        var order = new OrderData
        {
            CustomerName = parsed.BasicInfo.Customer ?? string.Empty,
            SaleDept = parsed.Sales.SalesDepartment ?? string.Empty,
            SaleManager = parsed.Sales.SalesManager ?? string.Empty,
            OrderDate = DateTime.Today.ToString("yyyy-MM-dd"),
            ContractAmount = ToText(parsed.BasicInfo.ContractTotalAmount),
            FlexibleNet = string.Join("；", parsed.Nets.Select(DescribeNet).Where(x => !string.IsNullOrWhiteSpace(x))),
            Spec = string.Join("、", parsed.Nets.SelectMany(n => n.Specs).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct()),
            BraidedSteelWireRope = string.Join("；", parsed.Nets.Select(DescribeRope).Where(x => !string.IsNullOrWhiteSpace(x))),
            UndergroundTransportMaxLength = parsed.UndergroundTransportSizes.FirstOrDefault()?.LengthM is { } transportLength ? ToText(transportLength) : string.Empty,
            Remark = BuildRemark(parsed)
        };

        foreach (var net in parsed.Nets)
        {
            var spec = FirstNonEmpty(net.Specs.FirstOrDefault(), net.ProductStandard, net.NetType);
            var term = FindPriceTerm(priceTerms, net.Specs);
            order.Items.Add(new OrderItem
            {
                ItemName = FirstNonEmpty(net.NetType, "网片"),
                Spec = spec,
                Unit = "平方米",
                Width = ToText(net.WidthM),
                Length = ToText(net.LengthM),
                LengthSegments = net.Segments?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                Quantity = Area(net.WidthM, net.LengthM),
                UnitPrice = ToText(term?.UnitPrice),
                Amount = ToText(term?.Amount),
                DeliveryDate = deliveryDate,
                ItemRemark = BuildNetRemark(net)
            });
        }

        foreach (var item in parsed.CarryItems)
        {
            order.Items.Add(new OrderItem
            {
                ItemName = item.Name ?? string.Empty,
                Spec = item.Spec ?? string.Empty,
                Unit = GuessUnit(item),
                Width = ToText(item.WidthM),
                Length = ToText(item.LengthM ?? item.TotalLengthM),
                Quantity = item.Pieces?.ToString(CultureInfo.InvariantCulture) ?? ToText(item.TotalLengthM),
                UnitPrice = ToText(item.UnitPrice),
                DeliveryDate = deliveryDate,
                ItemRemark = FirstNonEmpty(item.Detail, "携带/直接发货成品")
            });
        }

        foreach (var accessory in parsed.Accessories)
        {
            order.Items.Add(new OrderItem
            {
                ItemName = accessory.Name ?? string.Empty,
                Spec = accessory.Model ?? string.Empty,
                Unit = "件",
                Quantity = ToText(accessory.Quantity),
                UnitPrice = ToText(accessory.UnitPrice),
                DeliveryDate = deliveryDate,
                ItemRemark = "产品及配件"
            });
        }

        return order;
    }

    private static PriceTerm? FindPriceTerm(IReadOnlyList<PriceTerm> terms, IReadOnlyList<string> specs)
    {
        foreach (var spec in specs.Where(s => !string.IsNullOrWhiteSpace(s)))
        {
            var term = terms.FirstOrDefault(t => !string.IsNullOrWhiteSpace(t.Spec) && t.Spec.Contains(spec, StringComparison.OrdinalIgnoreCase));
            if (term != null) return term;
        }

        return terms.FirstOrDefault(t => t.UnitPrice.HasValue || t.Amount.HasValue);
    }

    private static string DescribeNet(NetItem net)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(net.NetType)) parts.Add(net.NetType!);
        if (net.LengthM.HasValue) parts.Add($"长{ToText(net.LengthM)}米");
        if (net.WidthM.HasValue) parts.Add($"宽{ToText(net.WidthM)}米");
        if (net.Segments.HasValue) parts.Add($"分{net.Segments}段");
        if (net.Specs.Count > 0) parts.Add($"规格{string.Join("、", net.Specs)}");
        return string.Join("，", parts);
    }

    private static string DescribeRope(NetItem net)
    {
        if (net.RopeProcess.Items.Count == 0 && !net.RopeProcess.ExposedEachEndM.HasValue) return string.Empty;
        var parts = net.RopeProcess.Items.Select(i => $"{i.Spec} 间距{ToText(i.SpacingM)}米 共{i.Count}根").ToList();
        if (net.RopeProcess.ExposedEachEndM.HasValue) parts.Add($"两端外露{ToText(net.RopeProcess.ExposedEachEndM)}米");
        return string.Join("，", parts);
    }

    private static string BuildNetRemark(NetItem net)
    {
        var sb = new StringBuilder();
        if (net.IsMixedWeaving && net.MixedWeaving.Count > 0)
        {
            sb.Append("混编：");
            sb.Append(string.Join("，", net.MixedWeaving.Select(x => $"{x.Spec}宽{ToText(x.WidthM)}米")));
        }

        var rope = DescribeRope(net);
        if (!string.IsNullOrWhiteSpace(rope))
        {
            if (sb.Length > 0) sb.Append("；");
            sb.Append("整编一般钢丝绳：").Append(rope);
        }

        return sb.ToString();
    }

    private static string BuildRemark(ParsedOrder parsed)
    {
        var warnings = parsed.Validation.Warnings.Concat(parsed.AiReview?.Warnings ?? []).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
        return warnings.Length == 0 ? "订单解析 Agent 自动生成" : "订单解析 Agent 自动生成；警告：" + string.Join("；", warnings);
    }

    private static string GuessUnit(CarryItem item)
    {
        if (item.Pieces.HasValue) return "件";
        if (item.TotalLengthM.HasValue) return "米";
        return "件";
    }

    private static string Area(decimal? width, decimal? length)
    {
        if (width is > 0 && length is > 0) return ToText(width.Value * length.Value);
        return string.Empty;
    }

    private static string ToText(decimal? value) => value.HasValue ? value.Value.ToString("0.####", CultureInfo.InvariantCulture) : string.Empty;
    private static string FirstNonEmpty(params string?[] values) => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))?.Trim() ?? string.Empty;
}

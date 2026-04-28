namespace AIWorkAssistant.Models.HkOrder;

public class OrderData
{
    public string CustomerName { get; set; } = "";
    public string SaleDept { get; set; } = "";
    public string SaleManager { get; set; } = "";
    public string OrderDate { get; set; } = "";
    public string ContractAmount { get; set; } = "";
    public string FlexibleNet { get; set; } = "";
    public string Spec { get; set; } = "";
    public string BraidedSteelWireRope { get; set; } = "";
    public string UndergroundTransportMaxLength { get; set; } = "";
    public string Remark { get; set; } = "";
    public List<OrderItem> Items { get; set; } = [];
}

public class OrderItem
{
    public string ItemCode { get; set; } = "";
    public string ItemName { get; set; } = "";
    public string Spec { get; set; } = "";
    public string Unit { get; set; } = "";
    public string Width { get; set; } = "";
    public string Length { get; set; } = "";
    public string LengthSegments { get; set; } = "";
    public string Quantity { get; set; } = "";
    public string UnitPrice { get; set; } = "";
    public string Amount { get; set; } = "";
    public string TaxRate { get; set; } = "13";
    public string DeliveryDate { get; set; } = "";
    public string ItemRemark { get; set; } = "";
}

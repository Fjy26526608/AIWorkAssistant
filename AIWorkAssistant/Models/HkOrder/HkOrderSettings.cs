namespace AIWorkAssistant.Models.HkOrder;

public class HkOrderSettings
{
    public string BrowserType { get; set; } = "Chromium";
    public bool Headless { get; set; } = false;
    public string TargetUrl { get; set; } = "";
    public string OrderPageUrl { get; set; } = "";
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public int DefaultTimeout { get; set; } = 30000;
    public int SlowMo { get; set; } = 0;
    public string OrderFolderPath { get; set; } = "";
    public string ApiKey { get; set; } = "";
    public string ApiBaseUrl { get; set; } = "";
    public string ModelName { get; set; } = "";
    public string SaleType { get; set; } = "煤矿";
}

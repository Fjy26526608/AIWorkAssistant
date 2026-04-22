namespace AIWorkAssistant.Services;

/// <summary>
/// 简易服务定位器（单例管理）
/// </summary>
public static class ServiceLocator
{
    public static AuthService Auth { get; } = new();
    public static ChatService Chat { get; } = new();
}

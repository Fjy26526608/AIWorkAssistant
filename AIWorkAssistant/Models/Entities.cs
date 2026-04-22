namespace AIWorkAssistant.Models;

/// <summary>
/// 用户
/// </summary>
public class User
{
    public int Id { get; set; }
    public string Username { get; set; } = string.Empty;
    public string PasswordHash { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string Role { get; set; } = "User"; // Admin / User
    public bool IsEnabled { get; set; } = true;
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    public List<UserAssistant> UserAssistants { get; set; } = new();
}

/// <summary>
/// AI 助手定义
/// </summary>
public class AiAssistant
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Icon { get; set; } = "🤖";
    public string SystemPrompt { get; set; } = string.Empty;
    public bool IsEnabled { get; set; } = true;
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    public List<UserAssistant> UserAssistants { get; set; } = new();
}

/// <summary>
/// 用户-助手关联（控制哪些用户能用哪些助手）
/// </summary>
public class UserAssistant
{
    public int UserId { get; set; }
    public User User { get; set; } = null!;
    public int AssistantId { get; set; }
    public AiAssistant Assistant { get; set; } = null!;
}

/// <summary>
/// 聊天消息
/// </summary>
public class ChatMessage
{
    public int Id { get; set; }
    public int UserId { get; set; }
    public int AssistantId { get; set; }
    public string Role { get; set; } = "user"; // user / assistant
    public string Content { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// 系统配置（键值对存储）
/// </summary>
public class AppSetting
{
    public string Key { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

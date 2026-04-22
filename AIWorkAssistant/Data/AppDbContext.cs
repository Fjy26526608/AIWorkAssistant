using Microsoft.EntityFrameworkCore;
using AIWorkAssistant.Models;

namespace AIWorkAssistant.Data;

public class AppDbContext : DbContext
{
    public DbSet<User> Users => Set<User>();
    public DbSet<AiAssistant> Assistants => Set<AiAssistant>();
    public DbSet<UserAssistant> UserAssistants => Set<UserAssistant>();
    public DbSet<ChatMessage> ChatMessages => Set<ChatMessage>();
    public DbSet<AppSetting> AppSettings => Set<AppSetting>();

    private readonly string _dbPath;

    public AppDbContext()
    {
        var appDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AIWorkAssistant");
        Directory.CreateDirectory(appDir);
        _dbPath = Path.Combine(appDir, "app.db");
    }

    protected override void OnConfiguring(DbContextOptionsBuilder options)
    {
        options.UseSqlite($"Data Source={_dbPath}");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // UserAssistant 多对多
        modelBuilder.Entity<UserAssistant>()
            .HasKey(ua => new { ua.UserId, ua.AssistantId });

        modelBuilder.Entity<UserAssistant>()
            .HasOne(ua => ua.User)
            .WithMany(u => u.UserAssistants)
            .HasForeignKey(ua => ua.UserId);

        modelBuilder.Entity<UserAssistant>()
            .HasOne(ua => ua.Assistant)
            .WithMany(a => a.UserAssistants)
            .HasForeignKey(ua => ua.AssistantId);

        // AppSetting 主键
        modelBuilder.Entity<AppSetting>()
            .HasKey(s => s.Key);

        // 默认管理员
        modelBuilder.Entity<User>().HasData(new User
        {
            Id = 1,
            Username = "admin",
            PasswordHash = HashPassword("admin123"),
            DisplayName = "管理员",
            Role = "Admin",
            IsEnabled = true,
            CreatedAt = DateTime.UtcNow
        });
    }

    public static string HashPassword(string password)
    {
        var bytes = System.Security.Cryptography.SHA256.HashData(
            System.Text.Encoding.UTF8.GetBytes(password));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }
}

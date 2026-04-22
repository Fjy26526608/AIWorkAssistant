using AIWorkAssistant.Data;
using AIWorkAssistant.Models;
using Microsoft.EntityFrameworkCore;

namespace AIWorkAssistant.Services;

public class AuthService
{
    public User? CurrentUser { get; private set; }

    public async Task<User?> LoginAsync(string username, string password)
    {
        await using var db = new AppDbContext();
        var hash = AppDbContext.HashPassword(password);
        var user = await db.Users.FirstOrDefaultAsync(u =>
            u.Username == username && u.PasswordHash == hash && u.IsEnabled);

        CurrentUser = user;
        return user;
    }

    public void Logout()
    {
        CurrentUser = null;
    }

    public bool IsLoggedIn => CurrentUser != null;
    public bool IsAdmin => CurrentUser?.Role == "Admin";
}

# AI Work Assistant

企业 AI 助手平台，基于 Avalonia UI 跨平台桌面应用。

## 功能

- 用户登录（账号密码认证）
- 根据用户权限加载不同的 AI 助手列表
- 选择 AI 助手后通过对话交互
- 管理员设置页面（AI 配置、用户管理、助手管理）
- SQLite 本地数据存储

## 技术栈

- .NET 10.0
- Avalonia UI 11.3
- Entity Framework Core + SQLite
- CommunityToolkit.Mvvm
- Anthropic Messages API

## 默认账号

- 用户名: `admin`
- 密码: `admin123`

## 运行

```bash
cd AIWorkAssistant
dotnet run
```

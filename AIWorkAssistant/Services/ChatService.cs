using System.Net.Http;
using System.Text;
using System.Text.Json;

namespace AIWorkAssistant.Services;

/// <summary>
/// AI 聊天服务（Anthropic Messages API 格式）
/// </summary>
public class ChatService
{
    private readonly HttpClient _httpClient;
    private string _baseUrl = "https://yxai.chat";
    private string _apiKey = "";
    private string _model = "glm-5.1";
    private int _maxTokens = 8192;

    public ChatService()
    {
        _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
    }

    public void Configure(string baseUrl, string apiKey, string model, int maxTokens = 8192)
    {
        _baseUrl = baseUrl.TrimEnd('/');
        _apiKey = apiKey;
        _model = model;
        _maxTokens = maxTokens;
    }

    public async Task<string> SendMessageAsync(string systemPrompt, List<(string role, string content)> messages)
    {
        var msgArray = messages.Select(m => new { role = m.role, content = m.content }).ToList();

        var requestBody = new Dictionary<string, object>
        {
            ["model"] = _model,
            ["max_tokens"] = _maxTokens,
            ["stream"] = false,
            ["messages"] = msgArray
        };

        if (!string.IsNullOrWhiteSpace(systemPrompt))
        {
            requestBody["system"] = systemPrompt;
        }

        var json = JsonSerializer.Serialize(requestBody);
        var request = new HttpRequestMessage(HttpMethod.Post, $"{_baseUrl}/v1/messages")
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey);

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.SendAsync(request);
        }
        catch (TaskCanceledException ex)
        {
            throw new TimeoutException("AI 请求超过 10 分钟仍未返回。请稍后重试，或检查网络和模型响应速度。", ex);
        }

        var responseText = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"AI 调用失败 ({response.StatusCode}): {responseText}");
        }

        using var doc = JsonDocument.Parse(responseText);
        var content = doc.RootElement.GetProperty("content");

        foreach (var block in content.EnumerateArray())
        {
            if (block.GetProperty("type").GetString() == "text")
            {
                return block.GetProperty("text").GetString() ?? "";
            }
        }

        throw new Exception("AI 响应中未找到文本内容");
    }
}

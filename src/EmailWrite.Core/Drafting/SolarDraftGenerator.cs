using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Upstage's Solar over its OpenAI-compatible chat-completions endpoint. Same prompt as every
/// other backend — <see cref="DraftPromptBuilder"/> is shared — so a draft from here can be scored
/// against the local model on the same axes with DraftProbe's <c>--repeat</c>.
///
/// No SDK: the request is one JSON object and the reply is one string, which is less code than
/// taking a dependency and leaves nothing to keep up to date. Like
/// <see cref="ClaudeDraftGenerator"/> the whole conversation is resent every turn, so there is no
/// KV cache to keep and <c>sessionId</c> is unused.
/// </summary>
public sealed class SolarDraftGenerator : IDraftGenerator
{
    // The endpoint and the model id are the two things most likely to have moved since this was
    // written, and getting either wrong is a 404/400 rather than a bad draft. Both are env-
    // overridable so a change costs a variable, not a rebuild, and the error path below prints
    // whatever the server actually said.
    private const string DefaultEndpoint = "https://api.upstage.ai/v1/chat/completions";
    private const string DefaultModel = "solar-pro2";

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(60) };

    /// <summary>The name <see cref="ApiKeys"/> files this backend's key under.</summary>
    public const string KeyName = "UPSTAGE_API_KEY";

    private readonly string _endpoint;
    private readonly string _model;

    public SolarDraftGenerator()
    {
        _endpoint = Environment.GetEnvironmentVariable("EMAILWRITE_SOLAR_ENDPOINT") ?? DefaultEndpoint;
        _model = Environment.GetEnvironmentVariable("EMAILWRITE_SOLAR_MODEL") ?? DefaultModel;
    }

    public async Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        int sessionId = 0,
        CancellationToken cancellationToken = default)
    {
        // Read per turn, not in the constructor: the popup keeps the generator it built the first
        // time this backend was picked, so a key entered afterwards would never reach a cached one.
        var apiKey = ApiKeys.Get(KeyName);
        if (apiKey == null)
        {
            return "Solar API 키가 없어요. 팝업 아래 입력칸에 키를 붙여넣고 저장해주세요.";
        }

        var (systemPrompt, examples) = DraftPromptBuilder.PromptFor(emailContext);
        var (openingTurn, rest) = DraftPromptBuilder.BuildTurns(emailContext, conversation);

        var messages = new List<object> { new { role = "system", content = systemPrompt } };
        messages.AddRange(examples.Select(t => new { role = Name(t.Role), content = t.Text }));
        messages.Add(new { role = "user", content = openingTurn });
        messages.AddRange(rest.Select(m => new { role = Name(m.Role), content = m.Text }));

        var request = new HttpRequestMessage(HttpMethod.Post, _endpoint)
        {
            Content = JsonContent.Create(new { model = _model, messages, max_tokens = 1024 }),
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        try
        {
            using var response = await Http.SendAsync(request, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                // The body, not just the status: a wrong model id and an expired key are both 4xx
                // and only the body tells them apart.
                return $"Solar API 오류 ({(int)response.StatusCode}): {Trim(body)}";
            }

            var text = JsonDocument.Parse(body).RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString();

            return string.IsNullOrWhiteSpace(text) ? "응답을 받지 못했어요." : text.Trim();
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException or KeyNotFoundException)
        {
            return $"Solar API 호출 중 오류가 발생했어요: {ex.Message}";
        }
    }

    private static string Name(ChatRole role) => role == ChatRole.User ? "user" : "assistant";

    private static string Trim(string body) => body.Length <= 300 ? body : body[..300] + "…";
}

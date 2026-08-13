using Anthropic;
using Anthropic.Models.Messages;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Calls the Anthropic Messages API to draft (and refine, on later turns) a
/// messenger-style reply based on the Outlook context injected as the system
/// prompt and the visible chat conversation.
/// </summary>
public sealed class ClaudeDraftGenerator : IDraftGenerator
{
    private const string ModelId = "claude-haiku-4-5";

    /// <summary>The name <see cref="ApiKeys"/> files this backend's key under.</summary>
    public const string KeyName = "ANTHROPIC_API_KEY";

    private AnthropicClient? _client;
    private string? _clientKey;

    public async Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        int sessionId = 0,
        CancellationToken cancellationToken = default)
    {
        // Built on the first turn rather than in the constructor, because the popup keeps the
        // generator it built when this backend was first picked: a key entered after that would
        // never reach a client that was decided at construction. Rebuilt only when the key itself
        // changed — the client holds its own HttpClient, and a new one per turn is how sockets get
        // exhausted, but a client still carrying the key the server just rejected is a 401 loop
        // with no way out of it from the UI.
        var apiKey = ApiKeys.Get(KeyName);
        if (apiKey == null)
        {
            return "Claude API 키가 없어요. 팝업 아래 입력칸에 키를 붙여넣고 저장해주세요.";
        }

        if (_client == null || _clientKey != apiKey)
        {
            _client = new AnthropicClient { ApiKey = apiKey };
            _clientKey = apiKey;
        }

        // Nothing is cached between calls here — the whole history is resent every turn — so
        // sessionId has nothing to do; it exists for LocalDraftGenerator's KV cache.
        var (systemPrompt, examples) = DraftPromptBuilder.PromptFor(emailContext);

        // The Outlook context is the opening user turn, so the conversation always starts
        // with a user message and stays correctly alternating on every follow-up turn.
        var messages = examples
            .Select(t => new MessageParam
            {
                Role = t.Role == ChatRole.User ? Role.User : Role.Assistant,
                Content = t.Text,
            })
            .ToList();
        var (openingTurn, rest) = DraftPromptBuilder.BuildTurns(emailContext, conversation);
        messages.Add(new MessageParam { Role = Role.User, Content = openingTurn });
        messages.AddRange(rest.Select(m => new MessageParam
        {
            Role = m.Role == ChatRole.User ? Role.User : Role.Assistant,
            Content = m.Text,
        }));

        Message response;
        try
        {
            response = await _client.Messages.Create(new MessageCreateParams
            {
                Model = ModelId,
                MaxTokens = 1024,
                System = systemPrompt,
                Messages = messages,
            });
        }
        catch (Exception ex)
        {
            return $"Claude API 호출 중 오류가 발생했어요: {ex.Message}";
        }

        var text = response.Content
            .Select(b => b.Value)
            .OfType<TextBlock>()
            .Select(t => t.Text)
            .FirstOrDefault();

        return string.IsNullOrWhiteSpace(text) ? "응답을 받지 못했어요." : text;
    }
}

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
    private readonly AnthropicClient? _client;

    public ClaudeDraftGenerator()
    {
        var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        if (!string.IsNullOrWhiteSpace(apiKey))
        {
            _client = new AnthropicClient { ApiKey = apiKey };
        }
    }

    public async Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        CancellationToken cancellationToken = default)
    {
        if (_client == null)
        {
            return "ANTHROPIC_API_KEY 환경변수가 설정되어 있지 않아요. 설정한 뒤 앱을 다시 시작해주세요.";
        }

        if (emailContext == null)
        {
            return "Outlook에서 열려있거나 선택된 메일을 찾지 못했어요. 메일을 연 뒤 다시 열어주세요.";
        }

        // The Outlook context is the opening user turn, so the conversation always starts
        // with a user message and stays correctly alternating on every follow-up turn.
        var messages = DraftPromptBuilder.ExampleTurns
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
                System = DraftPromptBuilder.SystemPrompt,
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

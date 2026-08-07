using System.Text;
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

        var messages = conversation
            .Select(m => new MessageParam
            {
                Role = m.Role == ChatRole.User ? Role.User : Role.Assistant,
                Content = m.Text,
            })
            .ToList();

        if (messages.Count == 0)
        {
            messages.Add(new MessageParam
            {
                Role = Role.User,
                Content = "위 메일에 대한 메신저 답장 초안을 작성해줘.",
            });
        }

        Message response;
        try
        {
            response = await _client.Messages.Create(new MessageCreateParams
            {
                Model = ModelId,
                MaxTokens = 1024,
                System = BuildSystemPrompt(emailContext),
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

    private static string BuildSystemPrompt(EmailContext context)
    {
        var sb = new StringBuilder();
        sb.AppendLine("당신은 사용자가 Outlook에서 받은 메일에 대해, 메신저(카카오톡/슬랙 등)로 보낼 답장 초안을 작성하는 어시스턴트입니다.");
        sb.AppendLine("이메일 형식(인사말, 서명 등)이 아니라 짧고 자연스러운 메신저 말투로 작성하세요.");
        sb.AppendLine("사용자가 이후 대화에서 톤이나 내용 수정을 요청하면 그에 맞춰 초안을 다듬어 다시 제시하세요.");
        sb.AppendLine();
        sb.AppendLine("## 현재 메일");
        AppendEmail(sb, context.Current);

        if (context.ThreadHistory.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("## 스레드 히스토리 (최근 순)");
            foreach (var message in context.ThreadHistory)
            {
                AppendEmail(sb, message);
            }
        }

        return sb.ToString();
    }

    private static void AppendEmail(StringBuilder sb, EmailMessage message)
    {
        sb.AppendLine($"- 제목: {message.Subject}");
        sb.AppendLine($"- 발신자: {message.SenderName}");
        sb.AppendLine($"- 수신 시각: {message.ReceivedTime:yyyy-MM-dd HH:mm}");
        if (message.AttachmentNames.Count > 0)
        {
            sb.AppendLine($"- 첨부파일: {string.Join(", ", message.AttachmentNames)}");
        }

        sb.AppendLine("- 본문:");
        sb.AppendLine(message.Body);
        sb.AppendLine();
    }
}

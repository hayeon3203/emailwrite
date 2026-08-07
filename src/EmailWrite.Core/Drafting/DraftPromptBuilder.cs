using System.Text;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Builds the prompt shared by every <see cref="IDraftGenerator"/> backend, regardless
/// of which model runs it. The role instructions go in the system prompt and the Outlook
/// context goes in the first user turn — keeping the mail body out of the system prompt
/// stops small local models from blurring "data" into "instructions" and echoing the
/// incoming mail back verbatim instead of answering it.
/// </summary>
internal static class DraftPromptBuilder
{
    public const string SystemPrompt = """
        당신은 사용자가 Outlook에서 받은 메일에 대해, 메신저(카카오톡/슬랙 등)로 보낼 답장 초안을 작성하는 어시스턴트입니다.
        초안은 메일을 '받은' 사용자가 '발신자에게' 보내는 답장입니다. 메일 문장을 그대로 되풀이하지 말고, 요청에 대한 사용자의 응답을 쓰세요.
        이메일 형식(인사말, 서명 등)이 아니라 짧고 자연스러운 메신저 말투로 작성하세요. 2~3문장이면 충분합니다.
        초안은 반드시 원본 메일과 같은 언어로 작성하세요. 한국어 메일에는 한국어로 답하세요.
        사용자가 이후 대화에서 톤이나 내용 수정을 요청하면 그에 맞춰 초안을 다듬어 다시 제시하세요.
        설명이나 머리말 없이 답장 본문만 출력하세요.
        """;

    /// <summary>
    /// A worked example, replayed as a completed user/assistant exchange ahead of the real
    /// one. Small local models need to *see* a correct draft — but only as a finished prior
    /// turn: the same example pasted into the system prompt gets copied out verbatim as the
    /// answer to the actual mail.
    /// </summary>
    public static readonly (string User, string Assistant) Example = (
        BuildContextTurn(new EmailContext(
            new EmailMessage(
                "계약서 검토 요청",
                "박지훈",
                "첨부한 계약서 검토 부탁드립니다. 수요일까지 회신 주세요.",
                new DateTime(2000, 1, 1),
                new[] { "계약서.pdf" }),
            Array.Empty<EmailMessage>())),
        "계약서 확인했습니다. 내일까지 검토해서 의견 정리해 드릴게요. 수요일 전에는 회신 드리겠습니다.");

    /// <summary>
    /// The opening user turn: the Outlook context as quoted data, then the instruction.
    /// The instruction goes last on purpose — small models follow the most recent text best.
    /// </summary>
    public static string BuildContextTurn(EmailContext context)
    {
        var sb = new StringBuilder();
        sb.AppendLine("## 답장할 메일");
        AppendEmail(sb, context.Current);

        if (context.ThreadHistory.Count > 0)
        {
            sb.AppendLine("## 스레드 히스토리 (최근 순, 참고용)");
            foreach (var message in context.ThreadHistory)
            {
                AppendEmail(sb, message);
            }
        }

        sb.Append($"위 메일을 보낸 '{context.Current.SenderName}'님에게 내가 보낼 메신저 답장 초안을 작성해줘.");
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

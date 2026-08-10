using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Stand-in for M5's Claude API call. Echoes back enough of the Outlook context
/// to prove the injection wiring works, without any network dependency.
/// </summary>
public sealed class PlaceholderDraftGenerator : IDraftGenerator
{
    public Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        int sessionId = 0,
        CancellationToken cancellationToken = default)
    {
        var reply = conversation.Count == 0 && emailContext != null
            ? $"\"{emailContext.Current.Subject}\" ({emailContext.Current.SenderName}) 메일 확인했어요. " +
              "답장 초안을 만들어드릴게요 — 실제 LLM 연동은 M5에서 붙습니다."
            : "아직 초안 생성 기능(M5)이 연결되기 전이라 이 대화에는 응답할 수 없어요.";

        return Task.FromResult(reply);
    }
}

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
        CancellationToken cancellationToken = default)
    {
        if (emailContext == null)
        {
            return Task.FromResult(
                "Outlook에서 열려있거나 선택된 메일을 찾지 못했어요. 메일을 연 뒤 다시 열어주세요.");
        }

        var reply = conversation.Count == 0
            ? $"\"{emailContext.Current.Subject}\" ({emailContext.Current.SenderName}) 메일 확인했어요. " +
              "답장 초안을 만들어드릴게요 — 실제 LLM 연동은 M5에서 붙습니다."
            : "아직 초안 생성 기능(M5)이 연결되기 전이라 이 대화에는 응답할 수 없어요.";

        return Task.FromResult(reply);
    }
}

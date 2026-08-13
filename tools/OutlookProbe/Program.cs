using EmailWrite.Core.Drafting;
using EmailWrite.Core.OutlookIntegration;

namespace OutlookProbe;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        if (!OutlookContextReader.TryGetCurrentEmailContext(out var context, out var error))
        {
            Console.WriteLine($"실패: {error}");
            return;
        }

        // 인사말이 여기서 나온다. 이름은 Outlook 프로필, 부서는 Exchange 전역 주소 목록(회사 서버)
        // 값이고 둘 다 앱이 만들어내는 게 아니라 받아오는 것이다 — 옛 소속이 찍히면 고칠 곳은
        // 여기가 아니라 그 주소록이다. 그걸 가르려면 실제로 뭐가 넘어오는지부터 보여야 한다.
        Console.WriteLine("=== 나 (인사말 재료) ===");
        Console.WriteLine($"이름(Outlook 프로필): {context!.UserName ?? "(못 읽음)"}");
        Console.WriteLine($"부서(회사 주소록): {context.UserOrg ?? "(못 읽음 — 비Exchange 계정이거나 접근 거부)"}");
        Console.WriteLine($"만들어지는 인사말: {DraftGreeting.For(context)?.Replace("\n", " / ") ?? "(없음)"}");
        Console.WriteLine();

        Console.WriteLine("=== 현재 메일 ===");
        Console.WriteLine($"제목: {context!.Current.Subject}");
        Console.WriteLine($"발신자: {context.Current.SenderName}");
        Console.WriteLine($"수신 시각: {context.Current.ReceivedTime}");
        Console.WriteLine($"첨부파일: {FormatAttachments(context.Current.AttachmentNames)}");
        Console.WriteLine($"본문 길이: {context.Current.Body.Length}자");
        Console.WriteLine($"본문 (앞 200자): {Truncate(context.Current.Body, 200)}");

        Console.WriteLine();
        Console.WriteLine($"=== 스레드 히스토리 ({context.ThreadHistory.Count}개) ===");
        foreach (var msg in context.ThreadHistory)
        {
            Console.WriteLine($"- [{msg.ReceivedTime}] {msg.SenderName}: {msg.Subject} (첨부: {FormatAttachments(msg.AttachmentNames)})");
            Console.WriteLine($"    본문 길이: {msg.Body.Length}자");
            Console.WriteLine($"    본문 (앞 100자): {Truncate(msg.Body, 100)}");
        }

        var promptChars =
            context.Current.Body.Length + context.ThreadHistory.Sum(m => m.Body.Length);
        Console.WriteLine();
        Console.WriteLine($"=== 프롬프트에 들어가는 본문 총합: {promptChars}자 ===");
    }

    private static string FormatAttachments(IReadOnlyList<string> names) =>
        names.Count == 0 ? "없음" : string.Join(", ", names);

    private static string Truncate(string s, int len) => s.Length <= len ? s : s[..len] + "...";
}

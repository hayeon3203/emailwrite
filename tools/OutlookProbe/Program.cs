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

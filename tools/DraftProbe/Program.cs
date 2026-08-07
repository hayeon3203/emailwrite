using System.Diagnostics;
using EmailWrite.Core.Drafting;
using EmailWrite.Core.OutlookIntegration;

// Exercises LocalDraftGenerator against a fixed fake EmailContext, so draft quality can be
// compared across models without Outlook running and without clicking through the popup.
// Swap models with the EMAILWRITE_MODEL env var (file must already be in the model cache):
//
//   $env:EMAILWRITE_MODEL="midm-2.0-mini-instruct-q4_k_m.gguf"
//   dotnet run --project tools\DraftProbe
//
// llama.cpp writes a few hundred lines of load/repack chatter to stderr; pipe it away with
//   ... 2>$null
Console.OutputEncoding = System.Text.Encoding.UTF8;

var context = new EmailContext(
    new EmailMessage(
        "[공유] 3분기 마케팅 캠페인 예산안 검토 요청",
        "김민수 팀장",
        "안녕하세요, 조하연님.\n첨부된 3분기 캠페인 예산안 초안 검토 부탁드립니다.\n특히 디지털 광고 비중을 40%까지 올리는 안에 대해 의견 주시면 좋겠습니다.\n금요일까지 회신 부탁드려요.\n감사합니다.",
        new DateTime(2026, 8, 5, 10, 12, 0),
        new[] { "Q3_캠페인_예산안_v2.xlsx" }),
    new[]
    {
        new EmailMessage(
            "RE: 2분기 결산 공유",
            "조하연",
            "확인했습니다. 2분기 디지털 채널 ROAS가 예상보다 좋았네요.",
            new DateTime(2026, 7, 28, 16, 40, 0),
            Array.Empty<string>()),
    });

Console.WriteLine($"모델: {Environment.GetEnvironmentVariable("EMAILWRITE_MODEL") ?? "(기본값)"}\n");

using var generator = new LocalDraftGenerator();

var sw = Stopwatch.StartNew();
Console.WriteLine("--- 1턴: 최초 초안 ---");
var draft = await generator.GenerateDraftAsync(context, Array.Empty<ChatMessage>());
Console.WriteLine(draft);
Console.WriteLine($"[{sw.Elapsed.TotalSeconds:F1}s — 모델 로드 포함]");

// Second turn checks that the model actually revises its own draft rather than
// re-answering the mail from scratch.
var conversation = new List<ChatMessage>
{
    ChatMessage.FromAssistant(draft),
    ChatMessage.FromUser("좀 더 짧고 캐주얼하게 바꿔줘."),
};

sw.Restart();
Console.WriteLine("\n--- 2턴: 톤 수정 (멀티턴) ---");
Console.WriteLine(await generator.GenerateDraftAsync(context, conversation));
Console.WriteLine($"[{sw.Elapsed.TotalSeconds:F1}s]");

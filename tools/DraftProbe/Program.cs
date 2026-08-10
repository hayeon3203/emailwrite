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
// Pass --outlook to run against the mail actually open in Outlook instead of the fake one, or
// --nomail to run the free-writing path, where there is no mail and only --intent drives the draft.
// The multi-turn quality bugs (wrong self-identification, email register creeping back in)
// only reproduce on real mail, whose body carries quoted headers and a signature block.
Console.OutputEncoding = System.Text.Encoding.UTF8;

if (args.Contains("--selftest"))
{
    // Real Outlook body: content, signature, then the quoted original with its headers.
    const string body =
        "몬드리안에서 영문버전 만들어주기로 했습니다.\n참고 부탁드립니다.\n\n김정희 드림\n" +
        "________________________________\n" +
        "보낸 사람: 박선영 <sunyoung.park@daolts.co.kr>\n받는 사람: 김정희\n\n원본 메일 본문";

    // Not Debug.Assert: that compiles away outside DEBUG builds, and a self-check that prints
    // "OK" without having checked anything is worse than no self-check.
    static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception($"selftest 실패: {message}");
    }

    var stripped = DraftPromptBuilder.StripQuotedReply(body);
    Check(stripped.EndsWith("김정희 드림"), $"인용부가 안 잘림: {stripped}");
    Check(!stripped.Contains("박선영"), "인용된 헤더의 다른 참여자 이름이 남아있음");

    // No marker at all -> body must survive untouched.
    Check(DraftPromptBuilder.StripQuotedReply("짧은 본문") == "짧은 본문", "마커 없는 본문이 잘림");

    // The earliest marker wins, even when a later one appears first in the marker list.
    Check(DraftPromptBuilder.StripQuotedReply("본문\nFrom: a\n________________________________\n") == "본문",
        "가장 앞선 마커에서 자르지 않음");

    Check(LocalDraftGenerator.StripChatMarkers("초안입니다.\n<|im_end|>") == "초안입니다.\n", "정지 토큰이 안 잘림");
    Check(LocalDraftGenerator.StripChatMarkers("초안입니다.") == "초안입니다.", "토큰 없는 출력이 잘림");
    Check(LocalDraftGenerator.StripChatMarkers("초안<|im_start|>assistant\n덧붙임") == "초안", "두 번째 마커에서 안 잘림");

    // 자유 작성 모드: 메일이 없으면 여는 턴에도 시스템 프롬프트에도 메일이 등장하면 안 됨.
    // 이게 깨지면 모델은 존재하지 않는 메일에 답장하려 들고, 없는 발신자 이름을 지어냄.
    var (freeTurn, freeRest) = DraftPromptBuilder.BuildTurns(
        null, new[] { ChatMessage.FromUser("내일 회의 30분 늦는다고") });
    Check(freeTurn.Contains("내일 회의 30분 늦는다고"), $"담을 내용이 여는 턴에서 빠짐: {freeTurn}");
    Check(!freeTurn.Contains("## 답장할 메일"), $"메일이 없는데 메일 섹션이 들어감: {freeTurn}");
    Check(freeRest.Count == 0, "첫 메시지가 여는 턴으로 접히지 않음");

    var mailCtx = new EmailContext(
        new EmailMessage("제목", "보낸이", "본문", new DateTime(2000, 1, 1), Array.Empty<string>()),
        Array.Empty<EmailMessage>());
    Check(DraftPromptBuilder.PromptFor(mailCtx).System.Contains("Outlook"), "메일이 있는데 자유 작성 프롬프트가 선택됨");
    Check(!DraftPromptBuilder.PromptFor(null).System.Contains("Outlook"), "메일이 없는데 메일 전제 프롬프트가 선택됨");
    Check(DraftPromptBuilder.PromptFor(null).Examples.Count > 0, "자유 작성 예시가 비어있음");

    Console.WriteLine("selftest OK");
    return;
}

var useOutlook = args.Contains("--outlook");

// Free-writing mode: no mail at all, the draft comes from --intent alone. This is what the popup
// does when Outlook has nothing open, and it exercises the other prompt set.
var noMail = args.Contains("--nomail");

// The second-turn request decides which failure shows up: "짧게" mostly behaves, while a tone
// change is what made the model prepend a preamble and slip back into email register.
var askIndex = Array.IndexOf(args, "--ask");
var secondTurnAsk = askIndex >= 0 && askIndex + 1 < args.Length
    ? args[askIndex + 1]
    : "조금 더 부드럽게 바꿔줘.";

// --draft pins the first-turn draft instead of generating one. Sampling makes turn 1 vary
// wildly, and the multi-turn failure only shows up when that draft came out terse ("네,
// 알겠습니다.") — with nothing to revise, the model re-answers the mail from scratch.
var draftIndex = Array.IndexOf(args, "--draft");
var pinnedDraft = draftIndex >= 0 && draftIndex + 1 < args.Length ? args[draftIndex + 1] : null;

// What the user wants the reply to say. The draft is built from this, not from the mail alone,
// so there is no first turn without it.
var intentIndex = Array.IndexOf(args, "--intent");
var intent = intentIndex >= 0 && intentIndex + 1 < args.Length
    ? args[intentIndex + 1]
    : "이번 주는 어려워서 다음 주 화요일에 회신하겠다고";

EmailContext? context = new EmailContext(
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

if (noMail)
{
    context = null;
    Console.WriteLine("메일 없음(자유 작성 모드)");
}
else if (useOutlook)
{
    // Outlook's COM objects are STA-only, and top-level statements can't carry [STAThread].
    EmailContext? live = null;
    string? error = null;
    var reader = new Thread(() => OutlookContextReader.TryGetCurrentEmailContext(out live, out error));
    reader.SetApartmentState(ApartmentState.STA);
    reader.Start();
    reader.Join();

    if (live == null)
    {
        Console.WriteLine($"Outlook 컨텍스트 읽기 실패: {error}");
        return;
    }

    context = live;
    Console.WriteLine($"실제 Outlook 메일 사용: \"{context.Current.Subject}\" ({context.Current.SenderName}, 본문 {context.Current.Body.Length}자)");
    if (args.Contains("--dump-body"))
    {
        // Only the head of the body survives DraftPromptBuilder's cap, so this is what the
        // model actually sees — including any quoted "받는 사람:" header it might mistake
        // for the user's own name.
        var forPrompt = DraftPromptBuilder.StripQuotedReply(context.Current.Body);
        Console.WriteLine($"--- 프롬프트에 들어가는 본문 ({context.Current.Body.Length}자 -> {forPrompt.Length}자) ---");
        Console.WriteLine(forPrompt[..Math.Min(1200, forPrompt.Length)]);
        Console.WriteLine("--- 끝 ---");
    }
}

// --dump-prompt prints what the model is handed, turn by turn, and stops. No weights are loaded:
// everything below the ChatML wrapper comes out of DraftPromptBuilder, so this shows the content
// exactly as the real run assembles it, without waiting on inference.
if (args.Contains("--dump-prompt"))
{
    var (dumpSystem, dumpExamples) = DraftPromptBuilder.PromptFor(context);
    var (dumpOpening, dumpRest) = DraftPromptBuilder.BuildTurns(
        context, new[] { ChatMessage.FromUser(intent) });

    Console.WriteLine($"===== [system] =====\n{dumpSystem}\n");
    var turn = 0;
    foreach (var (role, text) in dumpExamples)
    {
        Console.WriteLine($"===== [{(role == ChatRole.User ? "user" : "assistant")}] 1-shot 예시 {++turn} =====\n{text}\n");
    }

    Console.WriteLine($"===== [user] 실제 여는 턴 =====\n{dumpOpening}\n");
    foreach (var message in dumpRest)
    {
        Console.WriteLine($"===== [{(message.Role == ChatRole.User ? "user" : "assistant")}] 이후 턴 =====\n{message.Text}\n");
    }

    return;
}

Console.WriteLine($"모델: {Environment.GetEnvironmentVariable("EMAILWRITE_MODEL") ?? "(기본값)"}\n");

using var generator = new LocalDraftGenerator();

var conversation = new List<ChatMessage> { ChatMessage.FromUser(intent) };

// --repeat N generates the first-turn draft N times and stops. Sampling makes a single run say
// nothing about how often the model collapses into "네, 알겠습니다." — that needs a sample, and
// relaunching the process per sample would pay the model load every time. The session id changes
// each round so the KV cache is dropped and every draft starts from the same cold prompt.
var repeatIndex = Array.IndexOf(args, "--repeat");
if (repeatIndex >= 0 && repeatIndex + 1 < args.Length && int.TryParse(args[repeatIndex + 1], out var repeat))
{
    for (var i = 0; i < repeat; i++)
    {
        var sample = await generator.GenerateDraftAsync(context, conversation, sessionId: i + 100);
        Console.WriteLine($"{i + 1,3}. [{sample.Length,3}자] {sample.Replace("\n", " ")}");
    }

    return;
}

var sw = Stopwatch.StartNew();
Console.WriteLine($"--- 1턴: 최초 초안 (담을 내용: \"{intent}\") ---");
string draft;
if (pinnedDraft != null)
{
    draft = pinnedDraft;
    Console.WriteLine($"{draft}  [--draft로 고정, 생성 안 함]");
}
else
{
    draft = await generator.GenerateDraftAsync(context, conversation);
    Console.WriteLine(draft);
    Console.WriteLine($"[{sw.Elapsed.TotalSeconds:F1}s — 모델 로드 포함]");
}

// Second turn checks that the model actually revises its own draft rather than
// re-answering the mail from scratch.
conversation.Add(ChatMessage.FromAssistant(draft));
conversation.Add(ChatMessage.FromUser(secondTurnAsk));

sw.Restart();
Console.WriteLine("\n--- 2턴: 톤 수정 (멀티턴) ---");
var revised = await generator.GenerateDraftAsync(context, conversation);
Console.WriteLine(revised);
Console.WriteLine($"[{sw.Elapsed.TotalSeconds:F1}s]");

// Third turn exercises the KV cache continuation a second time in a row: if the bookkeeping
// drifts by even one token, this is where it shows up as an answer that ignores the request or
// talks about the wrong thing. "한 문장으로" is used because its effect is visible at a glance.
conversation.Add(ChatMessage.FromAssistant(revised));
conversation.Add(ChatMessage.FromUser("한 문장으로 짧게 줄여줘."));

sw.Restart();
Console.WriteLine("\n--- 3턴: 한 문장으로 축약 ---");
Console.WriteLine(await generator.GenerateDraftAsync(context, conversation));
Console.WriteLine($"[{sw.Elapsed.TotalSeconds:F1}s]");

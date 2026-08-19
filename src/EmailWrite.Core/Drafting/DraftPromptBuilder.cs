using System.Text;
using System.Text.RegularExpressions;
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
    // A real mail body is mostly signature, disclaimer and quoted history: the first real mail
    // tested against was 4,585 chars for four sentences of content. Uncapped, that blows past the local
    // model's 4k context (llama.cpp answers "NoKvSlot") and, when it does fit, spends minutes of
    // CPU prefill on boilerplate. The reply is always near the top, so keeping the head is enough.
    private const int MaxCurrentBodyChars = 1200;
    private const int MaxHistoryBodyChars = 300;

    /// <summary>
    /// Where Outlook starts quoting the message being replied to. Everything past this point is
    /// the older mail repeated verbatim, headers and all — and those headers are actively harmful,
    /// not merely wasteful: the model read "보낸 사람: 박선영" out of a quoted block and signed the
    /// draft as 박선영, who was another participant in the thread and not the user at all.
    ///
    /// Declared ahead of <see cref="ExampleTurns"/> on purpose: that field is initialised by
    /// calling into BuildContextTurn, and static initialisers run in declaration order, so
    /// putting this below it leaves the array null at exactly the moment it gets read.
    /// </summary>
    private static readonly string[] QuotedReplyMarkers =
    [
        "\n________________________________",   // Outlook's own separator above a quoted reply
        "\n-----Original Message-----",
        "\n-----원본 메시지-----",
        "\n보낸 사람:",
        "\nFrom:",
    ];

    /// <summary>
    /// Where the sender stops writing and their signature block starts. What is left after
    /// <see cref="QuotedReplyMarkers"/> have done their work is mostly this: one real mail ran
    /// 8,267 chars for two sentences of content, the rest being job title, company, phone, fax,
    /// address, a SharePoint URL and "Ideas spark change!!". The model reads all of it as material
    /// for the draft, and the prompt pays prefill time for every char of it.
    ///
    /// A Korean sign-off ends the content — everything below "최서윤 드림." is signature by
    /// convention. Capped at 20 chars before the sign-off word so a sentence that merely happens to
    /// end in 올림 doesn't take the rest of the mail with it.
    /// </summary>
    private static readonly Regex SignOffLine = new(
        @"^\s*.{0,20}(드림|올림|배상)\.?\s*$|^\s*(best regards|kind regards|regards|sincerely)[,.]?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// The contact lines of a signature, for the mails that carry one without a sign-off above it.
    /// Dropped line by line rather than cutting to the end, so a false positive costs one line
    /// instead of the remainder of the mail. The single-letter labels (M./F./E./A./T.) are what
    /// Korean corporate signatures use, and requiring a digit or an @ on the same line keeps them
    /// off an enumerated list item like "A. 첫 번째 안".
    /// </summary>
    private static readonly Regex ContactLine = new(
        @"^\s*(<?https?://|[MFETAHP]\.\s.*[\d@]|tel[.:]|fax|mobile|e-?mail\s*[:.])",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// A line carrying an Outlook-style link, which sits *inside* the line rather than starting it:
    /// <c>Daol Fusion Brochure(영문) 3.pdf &lt;https://…sharepoint.com/…&gt;</c>. <see cref="ContactLine"/>
    /// is anchored at the start and missed every one of these, and the drafts pasted the sender's
    /// own attachment URL back at them — 5 times in 20. §10 already put links outside the local
    /// model's scope, so nothing of value is lost by dropping the line.
    /// </summary>
    private static readonly Regex LinkLine = new(@"<https?://", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// The greeting a Korean business mail opens with. Kept short on purpose so it matches
    /// "안녕하세요, 하연님" but not a sentence that merely begins with a greeting.
    /// </summary>
    private static readonly Regex GreetingLine = new(@"^\s*안녕하세요[ ,.]*.{0,15}$", RegexOptions.Compiled);

    /// <summary>
    /// Which pair of prompts a turn runs under. There is no switch and no keyword: Outlook either
    /// had a mail or it did not, and that answer is already known before the model is asked
    /// anything. Two separate prompt sets rather than one prompt with conditional lines — at 2.3B
    /// the draft body is what degrades when the instructions grow, so the mail path keeps exactly
    /// the wording it has today.
    /// </summary>
    public static (string System, IReadOnlyList<(ChatRole Role, string Text)> Examples) PromptFor(
        EmailContext? context, string? kind = null)
    {
        // Candidate A, off by default. The classifier's answer picks the example set instead of
        // being written into the turn, so nothing here says which way the reply points — the shown
        // round is the only thing that carries it. Mail path only for now: the free-writing set has
        // a failure of its own (it gets copied out verbatim, Plans.md M8) that has to be settled
        // before its rounds are worth splitting by direction.
        if (ByKind)
        {
            // 공유·판정불가는 보여줄 라운드가 없으니 섞인 세트로 떨어진다. 그 세트도 성격 줄이
            // 없는 쪽이어야 한다 — 실제 턴은 이 후보에선 성격 줄을 절대 안 달기 때문에, 라벨
            // 달린 예시를 쓰면 모델이 실제 턴에서 처음 보는 형식이 된다.
            return context == null
                ? (FreeSystemPrompt, FreeExampleTurns)
                : (SystemPrompt, MailExamplesFor(kind) ?? ExampleTurns);
        }

        // The examples carry the 성격 field exactly when the real turn will. A field that first
        // appears in the real turn is one the model has never seen used (§9-1), and the reverse —
        // examples declaring a direction the real turn leaves blank — invites it to pick one.
        var withKind = kind != null;
        return context == null
            ? (FreeSystemPrompt, withKind ? FreeExampleTurnsWithKind : FreeExampleTurns)
            : (SystemPrompt, withKind ? ExampleTurnsWithKind : ExampleTurns);
    }

    /// <summary>
    /// Which way the reply points, decided before the draft prompt is built and handed in as data.
    /// §9(4) is the failure it exists for: "p.3 글자 굵기를 얇게 해주실 수 있는지" (a request the
    /// user is making) came back as "p.3의 글자 굵기는 더 얇게 수정했고" (a promise the user never
    /// made) — a sentence clean enough to copy out and send, saying the opposite of what was meant.
    /// 9-4 put one request-shaped example in and cut that from 18/20 to 3/20, but promise-shaped
    /// intents then flipped the other way (위험 0 → 5). No example set fixes both directions,
    /// because the model is guessing at something that can simply be established first.
    ///
    /// The wording carries a negation on purpose. "상대에게 부탁하는 내용" alone leaves the
    /// inversion available; naming the thing it must not become is what the failure actually is.
    /// </summary>
    private static readonly Dictionary<string, string> KindDescriptions = new()
    {
        ["요청"] = "상대에게 부탁하는 내용. 내가 이미 했다는 말이 아님",
        ["약속"] = "내가 하겠다고 하는 내용. 상대에게 부탁하는 게 아님",
        ["공유"] = "사실을 알리는 내용. 부탁도 약속도 아님",
    };

    /// <summary>
    /// The classifier's own prompt — deliberately the shortest thing in this file. It never sees the
    /// mail, the thread history or the signature block, only the one sentence the user typed. That
    /// is the whole reason a model which gets the direction wrong while drafting can still get it
    /// right here: §9 traced the inversion to material in the mail, and none of it is in this call.
    /// </summary>
    public const string KindSystemPrompt = """
        문장이 어떤 성격인지 판단해서 한 단어로만 답하세요. 설명하지 마세요.
        요청 = 상대방에게 무언가를 해달라고 부탁하는 것
        약속 = 내가 무언가를 하겠다고 하는 것
        공유 = 이미 있었던 사실을 알리기만 하는 것
        두 가지가 섞여 있으면, 상대방이 해야 할 일이 있는 쪽을 고르세요.
        """;

    public static string KindQuestion(string intent) => $"\"{intent}\"\n이 문장의 성격은?";

    /// <summary>
    /// Worked answers for the classifier, as completed turns — the same shape that fixed the draft
    /// prompt (§2: an example pasted into the system prompt gets copied out verbatim, the same
    /// example replayed as prior turns gets imitated). Without these the model answered in whatever
    /// register it felt like, and 9 of 20 runs came back as something ParseKind could not read.
    ///
    /// The last pair is the mixed case on purpose: real intents are usually "확인했다고, 그리고
    /// ~해달라고", and that is exactly where the model has to pick a side rather than hedge.
    /// </summary>
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> KindExampleTurns =
    [
        (ChatRole.User, KindQuestion("내일까지 검토해서 의견 정리해서 드린다고")),
        (ChatRole.Assistant, "약속"),
        (ChatRole.User, KindQuestion("지난번 견적서 최신본으로 다시 보내달라고")),
        (ChatRole.Assistant, "요청"),
        (ChatRole.User, KindQuestion("어제 배포는 문제없이 끝났다고")),
        (ChatRole.Assistant, "공유"),
        (ChatRole.User, KindQuestion("회의록 확인했다고, 발표 자료도 같이 공유해달라고")),
        (ChatRole.Assistant, "요청"),
    ];

    /// <summary>
    /// Reads the classifier's answer back. Anything that is not one of the three words yields null,
    /// and null means the draft prompt is built exactly as it is today — a wrong label points the
    /// draft the wrong way, which is worse than no label at all, so "모르겠음" has to be reachable.
    /// The earliest match wins, for the answers that ramble past the first word ("요청입니다. 왜냐하면").
    /// </summary>
    public static string? ParseKind(string answer)
    {
        string? found = null;
        var at = int.MaxValue;
        foreach (var kind in KindDescriptions.Keys)
        {
            var index = answer.IndexOf(kind, StringComparison.Ordinal);
            if (index >= 0 && index < at)
            {
                (at, found) = (index, kind);
            }
        }

        return found;
    }

    public const string SystemPrompt = """
        당신은 Outlook에서 사용자가 받은 메일에 대해 답장 초안을 작성하는 어시스턴트입니다.

        상황:
        - 사용자는 원본 메일의 수신자입니다.
        - 원본 메일 발신자가 사용자에게 요청하거나 문의했습니다.
        - 당신은 사용자가 원본 메일 발신자에게 보낼 답장을 작성합니다.
        - 즉, 답장 방향은 항상 “사용자 → 원본 메일 발신자”입니다.

        가장 중요한 금지 규칙:
        - 사용자가 입력한 “답장에 담을 내용”을 작업 완료 사실로 바꾸지 마세요.
        - 사용자가 요청문을 입력했다면, 그 요청을 상대방에게 전달하는 문장으로 작성하세요.
        - 원본 메일 발신자의 입장에서 쓰지 마세요.

        작성 규칙:
        - 사용자가 답장에 담으라고 한 내용만 자연스럽고 정중한 문장으로 옮기세요.
        - 받는 사람은 상사나 윗사람이라고 가정하고, 항상 정중한 존댓말로 작성하세요.
        - 이메일 형식(인사말, 서명 등)은 넣지 말고 메신저로 보낼 짧은 메시지로 작성하세요. 2~3문장이면 충분합니다.
        - 초안은 반드시 원본 메일과 같은 언어로 작성하세요. 한국어 메일에는 한국어로 답하세요.
        - 설명이나 머리말 없이 답장 본문만 출력하세요.
        """;

    /// <summary>
    /// The same job with the mail taken away: no sender to answer, no incoming text to avoid
    /// echoing, no original language to match. Every line of <see cref="SystemPrompt"/> that
    /// exists because of the mail is simply absent here rather than reworded, which makes this
    /// the easier of the two tasks for a small model — what is left is "turn these points into a
    /// message", with nothing to confuse it about who is speaking to whom.
    /// </summary>
    public const string FreeSystemPrompt = """
        당신은 사용자가 메신저(카카오톡/슬랙 등)로 보낼 메시지 초안을 대신 작성하는 어시스턴트입니다.
        사용자가 알려준 내용만 자연스러운 문장으로 옮기세요. 거기 없는 약속·일정·수치는 절대 지어내지 마세요.
        받는 사람은 상사나 윗사람입니다. 항상 정중한 존댓말로 쓰되, 이메일 형식(인사말, 서명 등)은 넣지 말고 메신저로 보낼 짧은 메시지로 작성하세요. 2~3문장이면 충분합니다.
        초안은 사용자가 쓴 언어로 작성하세요.
        사용자가 이후 대화에서 톤이나 내용 수정을 요청하면, 알겠다고 대답하지 말고 수정한 초안 전문을 다시 출력하세요.
        설명이나 머리말 없이 메시지 본문만 출력하세요.
        """;

    /// <summary>
    /// A worked exchange, replayed as completed turns ahead of the real one. Small local models
    /// need to *see* a correct draft — but only as finished prior turns: the same example pasted
    /// into the system prompt gets copied out verbatim as the answer to the actual mail.
    ///
    /// It runs to two rounds on purpose. With only the first pair, the model had never seen what
    /// a revision request is supposed to produce and treated it as conversation — answering "네,
    /// 물론이죠." instead of redrafting, or re-answering the mail from scratch in full email
    /// register, signature and all (and signing as whichever name it found in the quoted headers).
    /// The second pair shows the shape: rewrite the whole draft, no preamble, no acknowledgement.
    /// </summary>
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> ExampleTurns = BuildMailExamples(null);

    /// <summary>
    /// The same four (or six) turns with the 성격 field filled in. Both values get demonstrated
    /// rather than only the one the classifier returned: the point is that the field decides which
    /// way the draft points, and a single value cannot show that.
    /// </summary>
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> ExampleTurnsWithKind =
        BuildMailExamples(new[] { "약속", "요청" });

    /// <summary>The revision request every example set ends on. One string, so the two sets can't drift.</summary>
    private const string SofterAsk = "조금 더 부드럽게 바꿔줘.";

    /// <summary>
    /// A worked round where the user is promising to do something. Returned as its pieces rather
    /// than as turns so the same text serves both example sets — the default one, which shows both
    /// directions, and the kind-matched one below, which shows only this.
    /// </summary>
    private static (string User, string Draft, string Softer) PromiseRound(string? kind) => (
        BuildContextTurn(
            new EmailContext(
                new EmailMessage(
                    "계약서 검토 요청",
                    "박지훈",
                    "첨부한 계약서 검토 부탁드립니다. 수요일까지 회신 주세요.",
                    new DateTime(2000, 1, 1),
                    new[] { "계약서.pdf" }),
                Array.Empty<EmailMessage>(),
                // The examples carry the "(나)" mark too — a field that shows up for the first time
                // in the real turn is a field the model has never seen used. A stand-in name, since
                // these turns are fiction; the real one comes from the Outlook profile.
                UserName: "김수현"),
            // The example carries an intent too, so the model sees that the draft is built from
            // what the user asked for — not invented from the mail alone.
            "내일까지 검토해서 의견 정리해서 준다고, 수요일 전에는 회신한다고",
            kind),
        "계약서 확인했습니다. 내일까지 검토해서 의견 정리해 드리겠습니다. 수요일 전에는 회신드리겠습니다.",
        // Rewritten end to end, not just at the sentence endings: with a timid example — one
        // verb softened, everything else identical — the model matched that degree of change and
        // handed back the draft verbatim on most revision turns. The example has to show how far
        // a tone request is meant to move the text.
        "네, 계약서 잘 받았습니다. 내일까지 꼼꼼히 살펴보고 의견 정리해서 전달드리겠습니다. 늦어도 수요일 전에는 회신드리도록 하겠습니다.");

    /// <summary>
    /// The same round with the direction reversed: the user is asking the other party to do
    /// something. <paramref name="kind"/> fills the 성격 field for <see cref="ExampleTurnsWithKind"/>;
    /// the kind-matched set below leaves it null, because there the direction is carried by which
    /// round is shown rather than by a line of text.
    /// </summary>
    private static (string User, string Draft, string Softer) RequestRound(string? kind) => (
        BuildContextTurn(
            new EmailContext(
                new EmailMessage(
                    "회의록 초안 공유",
                    "이수진",
                    "어제 회의록 초안 공유드립니다. 확인 부탁드립니다.",
                    new DateTime(2000, 1, 1),
                    Array.Empty<string>()),
                Array.Empty<EmailMessage>(),
                UserName: "김수현"),
            "확인했다고, 발표 자료도 같이 공유해달라고",
            kind),
        "회의록 확인했습니다. 혹시 발표 자료도 같이 공유해주실 수 있을까요?",
        // Only reached by the kind-matched set — the default set's revision round belongs to the
        // promise draft. Moves as far as the promise one does, and stays a request the whole way:
        // a "softer" example that quietly turns into an offer would teach exactly the failure this
        // candidate exists to fix.
        "네, 회의록 잘 봤습니다. 정리해주셔서 감사합니다. 혹시 발표 자료도 함께 공유해주시면 큰 도움이 될 것 같습니다.");

    /// <summary>
    /// A second request round, so the kind-matched set can run two content rounds like the default
    /// one does. Candidate A's first cut showed only one, and the prompt it produced was two turns
    /// shorter than the measured baseline — CLOVA answered that by pasting in a subject line and
    /// copying sentences out of the thread, the same drift that showed up when the mail was removed
    /// entirely. Length is not a free variable here; matching it is the point of this round.
    ///
    /// Shape chosen deliberately: the other party already sent something, and the reply thanks them
    /// and asks for one more thing. That is what real intents look like, and it is where the
    /// direction has the most room to flip.
    /// </summary>
    private static (string User, string Draft) RequestRound2(string? kind) => (
        BuildContextTurn(
            new EmailContext(
                new EmailMessage(
                    "수정한 자료 전달드립니다",
                    "정민아",
                    "요청하신 부분 반영해서 수정한 파일 전달드립니다. 확인 부탁드립니다.",
                    new DateTime(2000, 1, 1),
                    Array.Empty<string>()),
                Array.Empty<EmailMessage>(),
                UserName: "김수현"),
            "수정본 확인했다고, 인쇄용 말고 배포용으로도 하나 만들어달라고",
            kind),
        "수정본 확인했습니다. 혹시 배포용으로도 하나 만들어주실 수 있을까요?");

    /// <summary><see cref="RequestRound2"/>'s counterpart, so both sets are the same length.</summary>
    private static (string User, string Draft) PromiseRound2(string? kind) => (
        BuildContextTurn(
            new EmailContext(
                new EmailMessage(
                    "월간 보고서 초안",
                    "한지우",
                    "월간 보고서 초안입니다. 수치 확인 부탁드립니다.",
                    new DateTime(2000, 1, 1),
                    Array.Empty<string>()),
                Array.Empty<EmailMessage>(),
                UserName: "김수현"),
            "오늘 안에 수치 확인해서 회신하겠다고",
            kind),
        "보고서 초안 확인했습니다. 오늘 안에 수치 확인해서 회신드리겠습니다.");

    /// <summary>
    /// Candidate A. The same worked rounds, but only the ones whose direction matches what the
    /// classifier said — a request-shaped intent sees request-shaped examples and nothing else.
    ///
    /// **This is deliberately not §11-2.** That fix put the direction in as a *sentence*
    /// ("- 성격: 상대에게 부탁하는 내용") and made drafts worse (깨끗 9 → 1), because a sentence with
    /// no subject let the model reach for the most visible request in the turn — the sender's.
    /// Here the direction is never stated. It is only demonstrated, which is the one thing that has
    /// worked in this file: every fix that stuck changed what the model is *shown*, and the three
    /// that told it something outright failed (§3).
    ///
    /// Worth trying because the classifier is the accurate part of this pipeline — 42/42 on the
    /// same sentences the draft call gets backwards 19 times in 20 (§11-1). The direction is known
    /// before the draft starts; only the way it was being spent was wrong.
    ///
    /// The revision pair stays in both sets. Without it the model answers "네, 알겠습니다." to a
    /// tone request instead of redrafting, and that failure has nothing to do with direction.
    /// EMAILWRITE_EX_ASK2 / EMAILWRITE_NO_EX_ASK do not apply here: those tune the ratio of the
    /// mixed set, and this set has no ratio.
    /// </summary>
    private static IReadOnlyList<(ChatRole Role, string Text)> BuildKindExamples(string kind)
    {
        var ask = kind == "요청";
        var (user, draft, softer) = ask ? RequestRound(null) : PromiseRound(null);
        // Same turn count and same running order as the default set — content, draft, tone request,
        // redraft, content, draft. Only the direction of the two content rounds differs, so an A/B
        // against the baseline moves exactly one thing.
        var (user2, draft2) = ask ? RequestRound2(null) : PromiseRound2(null);
        return
        [
            (ChatRole.User, user),
            (ChatRole.Assistant, draft),
            (ChatRole.User, SofterAsk),
            (ChatRole.Assistant, softer),
            (ChatRole.User, user2),
            (ChatRole.Assistant, draft2),
        ];
    }

    private static readonly IReadOnlyList<(ChatRole Role, string Text)> MailExamplesRequest = BuildKindExamples("요청");
    private static readonly IReadOnlyList<(ChatRole Role, string Text)> MailExamplesPromise = BuildKindExamples("약속");

    /// <summary>
    /// Which set candidate A hands over. 공유 and an unreadable answer both fall through to the
    /// mixed set — there is no 공유 round to show, and inventing one to fill the slot would be
    /// showing the model a shape nobody measured.
    /// </summary>
    private static IReadOnlyList<(ChatRole Role, string Text)>? MailExamplesFor(string? kind) => kind switch
    {
        "요청" => MailExamplesRequest,
        "약속" => MailExamplesPromise,
        _ => null,
    };

    private static IReadOnlyList<(ChatRole Role, string Text)> BuildMailExamples(string[]? kinds)
    {
        var promise = PromiseRound(kinds?[0]);
        var turns = new List<(ChatRole, string)>
        {
            (ChatRole.User, promise.User),
            (ChatRole.Assistant, promise.Draft),
            // Not "친근하게": the recipient is the user's superior, so a revision request that asks
            // for casual register would teach the model to leave the polite one behind. "부드럽게"
            // moves the text just as far without changing who it is being written to.
            (ChatRole.User, SofterAsk),
            (ChatRole.Assistant, promise.Softer),
        };

        // §9-4. Both rounds above are the same shape — a request comes in, the draft promises to
        // do it — and the model applied that shape to intents that are the other way round: asked
        // for "국문 버전도 만들어달라고" (please also make a Korean version) it wrote "만들어
        // 드리겠습니다" (I will make it) in 20 runs out of 20, inventing a commitment the user
        // could have copied out and sent. With this round added, 17 of 20 kept the request a
        // request.
        //
        // It is not free: the same 20 runs on a promise-shaped intent went the other way, five of
        // them flipping into "이번 주 바쁘시다고 하셨으니" — the model now has a shape where the
        // other party is the one doing things, and overapplies that instead. Across both intents
        // the dangerous class (meaning inverted, sentence clean enough to send) still drops 18/40
        // to 8/40, which is why this is on. EMAILWRITE_NO_EX_ASK=1 goes back to the two-round
        // prompt, 304 chars shorter (~21% of the prompt, and prefill is ~93% of a turn).
        if (!Flag("EMAILWRITE_NO_EX_ASK"))
        {
            var request = RequestRound(kinds?[1]);
            turns.Add((ChatRole.User, request.User));
            turns.Add((ChatRole.Assistant, request.Draft));
        }

        // 예시 비율은 지금 약속 2 : 요청 1이다. §9-4에서 요청 예시 하나가 0/20 -> 17/20을 만들었으니
        // 비율 자체가 움직이는 축일 수 있다. 이 쌍은 "상대가 이미 뭔가를 보내왔고, 거기에 더 요청한다"는
        // 지금 실패하는 의도와 같은 모양이다 — 받은 것에 감사하고 곧바로 다음 요청으로 넘어간다.
        if (Flag("EMAILWRITE_EX_ASK2"))
        {
            var (askUser2, askDraft2) = RequestRound2(null);
            turns.Add((ChatRole.User, askUser2));
            turns.Add((ChatRole.Assistant, askDraft2));
        }

        return turns.Select(t => (t.Item1, t.Item2)).ToList();
    }

    /// <summary>
    /// <see cref="ExampleTurns"/> for the no-mail case, kept to the same two rounds and the same
    /// second-round wording. The revision round is not optional here either: it is what stops the
    /// model from answering "네, 그렇게 해드릴게요." to a tone change instead of redrafting.
    /// </summary>
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> FreeExampleTurns = BuildFreeExamples(null);

    /// <summary><see cref="ExampleTurnsWithKind"/> for the no-mail path.</summary>
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> FreeExampleTurnsWithKind =
        BuildFreeExamples("요청");

    /// <summary>
    /// Experiment scaffolding, to be deleted once the free-writing candidates are settled: each
    /// switch is off unless its variable is "1", so the measured baseline is what runs by default
    /// and one candidate at a time can be turned on. Same idea as EMAILWRITE_MODEL/THREADS/TEMP.
    /// </summary>
    // defaultOn is for the candidates that won their A/B and became the shipped behaviour: unset
    // means on, and "0" is how they get turned back off for a comparison run.
    private static bool Flag(string name, bool defaultOn = false) =>
        Environment.GetEnvironmentVariable(name) is { } value ? value == "1" : defaultOn;

    /// <summary>Candidate A (§11-3), now the default: the direction is carried by which example
    /// rounds are shown, never by a 성격 line. EMAILWRITE_EX_BY_KIND=0 goes back to §11-2.</summary>
    private static bool ByKind => Flag("EMAILWRITE_EX_BY_KIND", defaultOn: true);

    private static IReadOnlyList<(ChatRole Role, string Text)> BuildFreeExamples(string? kind)
    {
        var turns = new List<(ChatRole, string)>
        {
            (ChatRole.User, BuildFreeTurn("내일 저녁 약속 못 갈 것 같다고, 다음 주 화요일로 미루자고", kind)),
            (ChatRole.Assistant, "죄송합니다, 내일 저녁 약속은 참석이 어려울 것 같습니다. 혹시 다음 주 화요일로 미뤄주실 수 있을까요?"),
            (ChatRole.User, "조금 더 부드럽게 바꿔줘."),
            (ChatRole.Assistant, "내일 저녁 약속 관련해서 말씀드립니다. 일정이 여의치 않아 참석이 어려울 것 같습니다. 괜찮으시다면 다음 주 화요일로 미뤄주시면 감사하겠습니다."),
        };

        // A/B knob for the copying bug on this path: with no mail, these four turns are most of
        // the prompt, and 6 of 20 request-shaped drafts came back as the second assistant turn
        // above, word for word (Experiments.md C-2). Rounds, not turns, because a round is the
        // unit that has ever been measured -- 2 removes the revision round (which section 2 says
        // is what stops a tone change from being answered conversationally), 0 removes everything.
        // Default 2 is the shipped baseline.
        var rounds = int.TryParse(
            Environment.GetEnvironmentVariable("EMAILWRITE_FREE_EX_ROUNDS"), out var r) ? r : 2;
        if (rounds < 2)
        {
            turns.RemoveRange(Math.Max(rounds, 0) * 2, turns.Count - Math.Max(rounds, 0) * 2);
        }

        // Candidate 2. Every turn above is one scenario -- calling off an appointment and moving
        // it -- and the drafts that failed worst applied exactly that move to content that did
        // not call for it ("회의 시작을 앞당겨주실 수 있을까요"). This pair is deliberately the
        // other shape, and the shape the failing test intent actually has: state something, then
        // ask for something, with both halves surviving into the draft.
        if (Flag("EMAILWRITE_EX2"))
        {
            turns.Add((ChatRole.User, BuildFreeTurn("자료 정리해서 공유 폴더에 올려뒀다고, 확인 후 피드백 달라고", kind)));
            turns.Add((ChatRole.Assistant, "요청하신 자료 정리해서 공유 폴더에 올려두었습니다. 확인해보시고 피드백 주시면 감사하겠습니다."));
        }

        return turns.Select(t => (t.Item1, t.Item2)).ToList();
    }

    /// <summary>
    /// Splits a conversation into the opening turn and whatever follows it. The user's first
    /// message says what the reply should contain, and it is folded into the same turn as the mail
    /// rather than sent as a message of its own: two user turns in a row is not a shape the chat
    /// templates (or the Anthropic API) expect, and the points to cover belong right next to the
    /// mail they are about. With no mail (<paramref name="context"/> null) that first message is
    /// all there is, and it becomes the opening turn on its own.
    /// </summary>
    public static (string OpeningTurn, IReadOnlyList<ChatMessage> Following) BuildTurns(
        EmailContext? context, IReadOnlyList<ChatMessage> conversation, string? kind = null)
    {
        // Candidate A spends the classifier's answer on the examples and nowhere else: the 성격 line
        // is dropped here rather than inside Points(), which the example sets are built from at type
        // load — a flag read there would bake whatever the environment said at first touch.
        if (ByKind)
            kind = null;

        var intent = conversation.Count > 0 && conversation[0].Role == ChatRole.User
            ? conversation[0].Text
            : null;
        var rest = intent != null ? conversation.Skip(1).ToList() : conversation;
        var opening = context == null
            ? BuildFreeTurn(intent, kind)
            : BuildContextTurn(context, intent, kind);
        return (opening, rest);
    }

    /// <summary>The opening turn with no mail behind it: just what the user wants said.</summary>
    private static string BuildFreeTurn(string? intent, string? kind = null) =>
        $"""
        ## 보낼 메시지에 담을 내용
        {Points(intent, kind)}

        위 내용으로 {FreeRecipient()}보낼 메신저 메시지 초안을 만들어줘. 담을 내용에 없는 약속이나 사실은 지어내지 마.
        """;

    /// <summary>
    /// The body of the 담을 내용 block: the points as typed, with the direction spelled out above
    /// them when one was established. Field lines rather than a sentence, because that is the shape
    /// the model is already reading everywhere else in the turn — and because every fix that stuck
    /// changed what the model is *shown*, while three that told it something outright failed (§3).
    /// </summary>
    private static string Points(string? intent, string? kind) =>
        kind == null || !KindDescriptions.TryGetValue(kind, out var description)
            ? intent ?? ""
            : $"- 성격: {description}\n- 내용: {intent}";

    /// <summary>
    /// Candidate 3. The mail path names who the draft is for ("'김민수 팀장'님에게 보낼"); the free
    /// path names nobody, so the only mention of a recipient is one line of the system prompt,
    /// far from the instruction the model follows best. A method rather than a field: this is
    /// called while <see cref="FreeExampleTurns"/> is being initialised, and a static field read
    /// there would be whatever it was before its own initialiser ran.
    /// </summary>
    private static string FreeRecipient() => Flag("EMAILWRITE_RCPT") ? "상사에게 " : "";

    /// <summary>
    /// The opening user turn: the Outlook context as quoted data, then the instruction.
    /// The instruction goes last on purpose — small models follow the most recent text best.
    /// </summary>
    public static string BuildContextTurn(EmailContext context, string? intent = null, string? kind = null)
    {
        // §9-1, off by default. Marking the user made nothing better on a mail with no thread
        // history — 20 runs each way, clean 17->13 on a request-shaped intent and 13->14 on a
        // promise-shaped one — and the half of this that targets the failure it was written for
        // (a draft copying a sentence out of the history and swapping the speaker) cannot be
        // measured until a mail that actually has history is the one open in Outlook.
        // Null switches every "(나)" mark below off at once, which is also what happens when
        // Outlook will not give up the profile name.
        var me = Flag("EMAILWRITE_WHOAMI") ? context.UserName : null;

        // §9-5, measured and reverted. Opening the turn with the points was meant to stop the model
        // writing from the mail instead of from them. It does the opposite: it hands the slot right
        // before the instruction to the mail body, and that slot is the one small models follow
        // best. On a mail whose body is itself a well-formed Korean business message, the easiest
        // continuation is to copy it — 20 runs gave 5 drafts signed as the sender and 5 that pasted
        // the sender's SharePoint URL, against 0 and 0 with the points back at the bottom.
        // EMAILWRITE_INTENT_FIRST=1 restores it for another look on a different mail.
        var intentFirst = Flag("EMAILWRITE_INTENT_FIRST");
        var hasIntent = !string.IsNullOrWhiteSpace(intent);

        var sb = new StringBuilder();

        void AppendIntent()
        {
            sb.AppendLine("## 답장에 담을 내용");
            sb.AppendLine(Points(intent, kind));
            sb.AppendLine();
        }

        if (intentFirst && hasIntent)
        {
            AppendIntent();
        }

        sb.AppendLine("## 답장할 메일");
        AppendEmail(sb, context.Current, MaxCurrentBodyChars, me, isCurrent: true);

        // An out-of-office reply is in the conversation but has nothing to do with the reply being
        // written: one thread carried 박선영's vacation auto-reply, and the draft that came out of
        // it read as if the user were the one away. Dropped before the header is written, or a
        // thread whose only other item is an auto-reply prints an empty history section.
        // §9(1)의 원래 용의자인데 한 번도 단독으로 안 쟀다. 히스토리는 "참고용"이지만, 이 메일에서는
        // 사용자 자신의 지난 메시지(p.8/p.22 수정 요청)를 통째로 실어와서 초안이 그걸 되풀이했다.
        var history = Flag("EMAILWRITE_NO_HISTORY")
            ? new List<EmailMessage>()
            : context.ThreadHistory.Where(m => !IsAutoReply(m)).ToList();
        if (history.Count > 0)
        {
            sb.AppendLine("## 스레드 히스토리 (최근 순, 참고용)");
            foreach (var message in history)
            {
                AppendEmail(sb, message, MaxHistoryBodyChars, me);
            }
        }

        if (!hasIntent)
        {
            sb.Append($"위 메일을 보낸 '{context.Current.SenderName}'님에게 내가 보낼 메신저 답장 초안을 작성해줘.");
            return sb.ToString();
        }

        if (intentFirst)
        {
            // "위 내용"은 메일을 가리키게 되므로 '담을 내용'을 이름으로 부른다.
            sb.Append($"맨 위 '답장에 담을 내용'을 '{context.Current.SenderName}'님에게 보낼 메신저 답장 초안으로 " +
                      "만들어줘. 담을 내용에 없는 약속이나 사실은 지어내지 마.");
            return sb.ToString();
        }

        AppendIntent();
        sb.Append($"위 내용을 '{context.Current.SenderName}'님에게 보낼 메신저 답장 초안으로 만들어줘. " +
                  "담을 내용에 없는 약속이나 사실은 지어내지 마.");
        return sb.ToString();
    }

    private static void AppendEmail(
        StringBuilder sb, EmailMessage message, int maxBodyChars, string? me = null, bool isCurrent = false)
    {
        sb.AppendLine($"- 제목: {message.Subject}");
        sb.AppendLine($"- 발신자: {message.SenderName}{(IsMe(message.SenderName, me) ? " (나)" : "")}");
        // Data, not another instruction. Every line the mail path gained by *telling* the model
        // something has failed (§3: three of them killed the output outright), while every fix that
        // stuck changed what the model is shown — so who the user is arrives as one more field of
        // the mail, in the shape the model is already reading.
        if (isCurrent && me != null)
        {
            sb.AppendLine($"- 받는 사람: {me} (나)");
        }

        sb.AppendLine($"- 수신 시각: {message.ReceivedTime:yyyy-MM-dd HH:mm}");
        if (message.AttachmentNames.Count > 0)
        {
            sb.AppendLine($"- 첨부파일: {string.Join(", ", message.AttachmentNames)}");
        }

        // 본문을 아예 빼보는 칸. 사용자가 담을 내용을 이미 말해줬는데 본문이 무엇을 보태는지는
        // 검증된 적이 없고, 오늘 데이터는 본문에 남은 유일한 신호(발신자의 요청)가 초안을 약속으로
        // 끌고 간다고 말한다. 제목·발신자는 남으므로 누구에게 쓰는지는 그대로 안다.
        var bodyCap = int.TryParse(Environment.GetEnvironmentVariable("EMAILWRITE_BODY"), out var cap)
            ? cap
            : maxBodyChars;
        if (bodyCap > 0)
        {
            sb.AppendLine("- 본문:");
            sb.AppendLine(Truncate(CleanBody(message.Body, message.SenderName), bodyCap));
        }
        sb.AppendLine();
    }

    /// <summary>
    /// The mail body with everything that is not the sender's message taken out — the quoted
    /// original first, then the signature. Truncation is what used to stand in for this, and it is
    /// a poor substitute: the cap keeps the first 1,200 chars whether they are content or a fax
    /// number, and history entries were being cut mid-signature at 300.
    /// </summary>
    internal static string CleanBody(string body, string? senderName = null) =>
        Environment.GetEnvironmentVariable("EMAILWRITE_KEEPSIG") == "1"
            ? StripQuotedReply(body)
            : StripGreeting(StripSignature(StripQuotedReply(body)), senderName);

    /// <summary>
    /// The two lines a Korean business mail opens with — "안녕하세요, 하연님" and "다올TS 마케팅팀
    /// 최서윤입니다." — which are the sender introducing themselves and say nothing about what the
    /// reply should contain.
    ///
    /// They are also, measurably, the thing the model copies when it writes the draft as the sender.
    /// A thread where the same person wrote twice puts that exact pair in the prompt twice, making
    /// it the most repeated well-formed text in the turn, and it is already in the shape of a
    /// message opening. §9-2 cut the tail of a mail for the same reason; this is the head.
    ///
    /// Only the run of greeting/intro lines at the very top goes: the walk stops at the first line
    /// that is neither, so a "…입니다" sentence further down is content and stays. The self-intro is
    /// matched against the sender's own name rather than a general pattern, so "전달드립니다" and
    /// other content ending the same way cannot be caught by it.
    /// </summary>
    internal static string StripGreeting(string body, string? senderName)
    {
        var lines = body.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                continue;
            }

            if (!GreetingLine.IsMatch(lines[i]) && !IsSelfIntro(lines[i], senderName))
            {
                break;
            }

            lines[i] = "";
        }

        return string.Join('\n', lines).Trim('\n').TrimEnd();
    }

    /// <summary>"다올TS 마케팅팀 최서윤입니다." — the sender naming themselves.</summary>
    private static bool IsSelfIntro(string line, string? senderName) =>
        !string.IsNullOrWhiteSpace(senderName)
        && line.Length < 40
        && line.Contains(senderName.Trim() + "입니다", StringComparison.Ordinal);

    /// <summary>
    /// Whether this sender is the user. Outlook writes the same display name into both places, so
    /// an exact match is enough — and a miss only costs the "(나)" mark, never a wrong one.
    /// </summary>
    private static bool IsMe(string senderName, string? me) =>
        me != null && string.Equals(senderName.Trim(), me.Trim(), StringComparison.OrdinalIgnoreCase);

    /// <summary>Out-of-office replies, which Outlook prefixes on the subject.</summary>
    private static bool IsAutoReply(EmailMessage message) =>
        message.Subject.StartsWith("자동 회신:", StringComparison.Ordinal) ||
        message.Subject.StartsWith("Automatic reply:", StringComparison.OrdinalIgnoreCase);

    private static string Truncate(string body, int maxChars) =>
        body.Length <= maxChars ? body : body[..maxChars] + "\n...(이하 생략)";

    internal static string StripQuotedReply(string body)
    {
        var cut = body.Length;
        foreach (var marker in QuotedReplyMarkers)
        {
            var index = body.IndexOf(marker, StringComparison.Ordinal);
            if (index >= 0 && index < cut)
            {
                cut = index;
            }
        }

        return body[..cut].TrimEnd();
    }

    internal static string StripSignature(string body)
    {
        var kept = new List<string>();
        foreach (var line in body.Split('\n'))
        {
            if (SignOffLine.IsMatch(line))
            {
                break;
            }

            if (!ContactLine.IsMatch(line) && !LinkLine.IsMatch(line))
            {
                kept.Add(line);
            }
        }

        return string.Join('\n', kept).TrimEnd();
    }
}

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
    /// Which pair of prompts a turn runs under. There is no switch and no keyword: Outlook either
    /// had a mail or it did not, and that answer is already known before the model is asked
    /// anything. Two separate prompt sets rather than one prompt with conditional lines — at 2.3B
    /// the draft body is what degrades when the instructions grow, so the mail path keeps exactly
    /// the wording it has today.
    /// </summary>
    public static (string System, IReadOnlyList<(ChatRole Role, string Text)> Examples) PromptFor(
        EmailContext? context) =>
        context == null ? (FreeSystemPrompt, FreeExampleTurns) : (SystemPrompt, ExampleTurns);

    public const string SystemPrompt = """
        당신은 사용자가 Outlook에서 받은 메일에 대해, 메신저(카카오톡/슬랙 등)로 보낼 답장 초안을 작성하는 어시스턴트입니다.
        초안은 메일을 '받은' 사용자가 '발신자에게' 보내는 답장입니다. 메일 문장을 그대로 되풀이하지 말고, 요청에 대한 사용자의 응답을 쓰세요.
        사용자가 "답장에 담을 내용"을 알려주면, 그 내용만 자연스러운 문장으로 옮기세요. 거기 없는 약속·일정·수치는 절대 지어내지 마세요.
        받는 사람은 상사나 윗사람입니다. 항상 정중한 존댓말로 쓰되, 이메일 형식(인사말, 서명 등)은 넣지 말고 메신저로 보낼 짧은 메시지로 작성하세요. 2~3문장이면 충분합니다.
        초안은 반드시 원본 메일과 같은 언어로 작성하세요. 한국어 메일에는 한국어로 답하세요.
        사용자가 이후 대화에서 톤이나 내용 수정을 요청하면, 알겠다고 대답하지 말고 수정한 초안 전문을 다시 출력하세요.
        설명이나 머리말 없이 답장 본문만 출력하세요.
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
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> ExampleTurns = new[]
    {
        (ChatRole.User, BuildContextTurn(
            new EmailContext(
                new EmailMessage(
                    "계약서 검토 요청",
                    "박지훈",
                    "첨부한 계약서 검토 부탁드립니다. 수요일까지 회신 주세요.",
                    new DateTime(2000, 1, 1),
                    new[] { "계약서.pdf" }),
                Array.Empty<EmailMessage>()),
            // The example carries an intent too, so the model sees that the draft is built from
            // what the user asked for — not invented from the mail alone.
            "내일까지 검토해서 의견 정리해서 준다고, 수요일 전에는 회신한다고")),
        (ChatRole.Assistant, "계약서 확인했습니다. 내일까지 검토해서 의견 정리해 드리겠습니다. 수요일 전에는 회신드리겠습니다."),
        // Not "친근하게": the recipient is the user's superior, so a revision request that asks
        // for casual register would teach the model to leave the polite one behind. "부드럽게"
        // moves the text just as far without changing who it is being written to.
        (ChatRole.User, "조금 더 부드럽게 바꿔줘."),
        // Rewritten end to end, not just at the sentence endings: with a timid example — one
        // verb softened, everything else identical — the model matched that degree of change and
        // handed back the draft verbatim on most revision turns. The example has to show how far
        // a tone request is meant to move the text.
        (ChatRole.Assistant, "네, 계약서 잘 받았습니다. 내일까지 꼼꼼히 살펴보고 의견 정리해서 전달드리겠습니다. 늦어도 수요일 전에는 회신드리도록 하겠습니다."),
    };

    /// <summary>
    /// <see cref="ExampleTurns"/> for the no-mail case, kept to the same two rounds and the same
    /// second-round wording. The revision round is not optional here either: it is what stops the
    /// model from answering "네, 그렇게 해드릴게요." to a tone change instead of redrafting.
    /// </summary>
    public static readonly IReadOnlyList<(ChatRole Role, string Text)> FreeExampleTurns = BuildFreeExamples();

    /// <summary>
    /// Experiment scaffolding, to be deleted once the free-writing candidates are settled: each
    /// switch is off unless its variable is "1", so the measured baseline is what runs by default
    /// and one candidate at a time can be turned on. Same idea as EMAILWRITE_MODEL/THREADS/TEMP.
    /// </summary>
    private static bool Flag(string name) => Environment.GetEnvironmentVariable(name) == "1";

    private static IReadOnlyList<(ChatRole Role, string Text)> BuildFreeExamples()
    {
        var turns = new List<(ChatRole, string)>
        {
            (ChatRole.User, BuildFreeTurn("내일 저녁 약속 못 갈 것 같다고, 다음 주 화요일로 미루자고")),
            (ChatRole.Assistant, "죄송합니다, 내일 저녁 약속은 참석이 어려울 것 같습니다. 혹시 다음 주 화요일로 미뤄주실 수 있을까요?"),
            (ChatRole.User, "조금 더 부드럽게 바꿔줘."),
            (ChatRole.Assistant, "내일 저녁 약속 관련해서 말씀드립니다. 일정이 여의치 않아 참석이 어려울 것 같습니다. 괜찮으시다면 다음 주 화요일로 미뤄주시면 감사하겠습니다."),
        };

        // Candidate 2. Every turn above is one scenario -- calling off an appointment and moving
        // it -- and the drafts that failed worst applied exactly that move to content that did
        // not call for it ("회의 시작을 앞당겨주실 수 있을까요"). This pair is deliberately the
        // other shape, and the shape the failing test intent actually has: state something, then
        // ask for something, with both halves surviving into the draft.
        if (Flag("EMAILWRITE_EX2"))
        {
            turns.Add((ChatRole.User, BuildFreeTurn("자료 정리해서 공유 폴더에 올려뒀다고, 확인 후 피드백 달라고")));
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
        EmailContext? context, IReadOnlyList<ChatMessage> conversation)
    {
        var intent = conversation.Count > 0 && conversation[0].Role == ChatRole.User
            ? conversation[0].Text
            : null;
        var rest = intent != null ? conversation.Skip(1).ToList() : conversation;
        var opening = context == null ? BuildFreeTurn(intent) : BuildContextTurn(context, intent);
        return (opening, rest);
    }

    /// <summary>The opening turn with no mail behind it: just what the user wants said.</summary>
    private static string BuildFreeTurn(string? intent) =>
        $"""
        ## 보낼 메시지에 담을 내용
        {intent}

        위 내용으로 {FreeRecipient()}보낼 메신저 메시지 초안을 만들어줘. 담을 내용에 없는 약속이나 사실은 지어내지 마.
        """;

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
    public static string BuildContextTurn(EmailContext context, string? intent = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("## 답장할 메일");
        AppendEmail(sb, context.Current, MaxCurrentBodyChars);

        if (context.ThreadHistory.Count > 0)
        {
            sb.AppendLine("## 스레드 히스토리 (최근 순, 참고용)");
            foreach (var message in context.ThreadHistory)
            {
                AppendEmail(sb, message, MaxHistoryBodyChars);
            }
        }

        if (string.IsNullOrWhiteSpace(intent))
        {
            sb.Append($"위 메일을 보낸 '{context.Current.SenderName}'님에게 내가 보낼 메신저 답장 초안을 작성해줘.");
            return sb.ToString();
        }

        sb.AppendLine("## 답장에 담을 내용");
        sb.AppendLine(intent);
        sb.AppendLine();
        sb.Append($"위 내용을 '{context.Current.SenderName}'님에게 보낼 메신저 답장 초안으로 만들어줘. " +
                  "담을 내용에 없는 약속이나 사실은 지어내지 마.");
        return sb.ToString();
    }

    private static void AppendEmail(StringBuilder sb, EmailMessage message, int maxBodyChars)
    {
        sb.AppendLine($"- 제목: {message.Subject}");
        sb.AppendLine($"- 발신자: {message.SenderName}");
        sb.AppendLine($"- 수신 시각: {message.ReceivedTime:yyyy-MM-dd HH:mm}");
        if (message.AttachmentNames.Count > 0)
        {
            sb.AppendLine($"- 첨부파일: {string.Join(", ", message.AttachmentNames)}");
        }

        sb.AppendLine("- 본문:");
        sb.AppendLine(Truncate(StripQuotedReply(message.Body), maxBodyChars));
        sb.AppendLine();
    }

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
}

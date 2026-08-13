using System.Text.RegularExpressions;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// The greeting and self-introduction pasted above a draft — "안녕하세요, 김민수 팀장님." /
/// "다올TS 마케팅팀 조하연입니다."
///
/// Built here rather than asked of the model, and that is the whole point of the class. §11-3
/// measured what happens when the system prompt stops forbidding email format: the model wants to
/// introduce itself, the only self-introduction anywhere in its context is the *sender's*, and it
/// copies that one — 20 runs went from 1 draft written in the sender's name to 7, and clean drafts
/// fell 9 → 4. §9-2 then went further and deliberately *deletes* the sender's greeting from the
/// mail before the model sees it, for exactly that reason. So there is no model-side way to ask
/// for this that does not walk back into the failure the prompt was fixed to avoid.
///
/// Every field here is already known before the model runs — who the mail is from, who the user
/// is, which department they are in. None of it needs generating, so none of it is generated.
/// </summary>
public static class DraftGreeting
{
    /// <summary>
    /// A greeting sentence the model wrote anyway, at the very start of a draft. The system prompt
    /// forbids it, so this is a backstop rather than the normal case — but when it does happen the
    /// user sees "안녕하세요" twice, which is the one failure mode this feature can introduce.
    ///
    /// The greeting has to *end* within 15 characters (a period or a line break) to count. Without
    /// that bound, "안녕하세요, 지난번에 말씀 주신 건으로 연락드립니다." — a greeting-shaped opening to
    /// a real sentence — loses its first fifteen characters and the draft is left mid-word.
    /// </summary>
    private static readonly Regex LeadingGreeting = new(@"^\s*안녕하세요[^.\n]{0,15}?[.\n]\s*", RegexOptions.Compiled);

    /// <summary>
    /// The block that goes above the draft, or null when it cannot be built honestly — no mail (so
    /// no one to greet), or Outlook would not say who the user is. A half-built greeting
    /// ("안녕하세요, 님.") is worse than none.
    /// </summary>
    public static string? For(EmailContext? context)
    {
        var recipient = context?.Current.SenderName?.Trim();
        if (string.IsNullOrEmpty(recipient))
        {
            return null;
        }

        var lines = new List<string> { $"안녕하세요, {recipient}님." };

        // The department alone cannot introduce anyone, so it only ever rides along with the name.
        var me = context!.UserName?.Trim();
        if (!string.IsNullOrEmpty(me))
        {
            var org = context.UserOrg?.Trim();
            lines.Add(string.IsNullOrEmpty(org) ? $"{me}입니다." : $"{org} {me}입니다.");
        }

        return string.Join('\n', lines);
    }

    /// <summary>
    /// The draft with its own opening greeting and self-introduction removed, so the block from
    /// <see cref="For"/> does not land on top of one. The self-introduction is matched against the
    /// user's actual name rather than a general "…입니다" pattern — "전달드립니다" and plenty of real
    /// content end the same way, and cutting content is a worse failure than a repeated line.
    /// </summary>
    public static string StripLeadingGreeting(string draft, string? userName)
    {
        draft = LeadingGreeting.Replace(draft, string.Empty, 1);

        var me = userName?.Trim();
        if (!string.IsNullOrEmpty(me))
        {
            var intro = new Regex($@"^\s*[^.\n]{{0,30}}?{Regex.Escape(me)}입니다[.\n]\s*");
            draft = intro.Replace(draft, string.Empty, 1);
        }

        return draft.TrimStart();
    }

    /// <summary>
    /// What the bubble shows and the 복사 button copies: greeting, blank line, draft. The model is
    /// never shown this — the popup keeps the raw draft as the conversation turn — so a following
    /// tone request revises the draft rather than the letterhead.
    /// </summary>
    public static string Apply(string draft, EmailContext? context)
    {
        var greeting = For(context);
        return greeting == null
            ? draft
            : $"{greeting}\n\n{StripLeadingGreeting(draft, context?.UserName)}";
    }
}

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Splits a model turn into "still talking" and "here is the reply to send". The model writes a
/// draft by wrapping it in &lt;draft&gt; tags, which is this codebase's stand-in for a tool call:
/// LLamaSharp has no function-calling support and the local 2.3B model could not be trusted with
/// it anyway, but it can manage a delimiter. Anthropic's real tool use would plug in at this same
/// seam — everything downstream only cares about the two fields below.
/// </summary>
public readonly record struct DraftResponse(bool IsDraft, string Text)
{
    private const string OpenTag = "<draft>";
    private const string CloseTag = "</draft>";

    public static DraftResponse Parse(string raw)
    {
        var start = raw.IndexOf(OpenTag, StringComparison.OrdinalIgnoreCase);
        if (start < 0)
        {
            return new DraftResponse(false, raw.Trim());
        }

        var contentStart = start + OpenTag.Length;
        var end = raw.IndexOf(CloseTag, contentStart, StringComparison.OrdinalIgnoreCase);

        // An unterminated tag means generation was cut off at the token limit. The draft is still
        // the most useful thing on offer, so take what there is rather than showing the user a
        // reply with a stray "<draft>" in front of it.
        var text = end < 0 ? raw[contentStart..] : raw[contentStart..end];
        return new DraftResponse(true, text.Trim());
    }
}

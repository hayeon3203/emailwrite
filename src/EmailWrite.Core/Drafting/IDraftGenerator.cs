using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Produces the assistant's next chat message given the Outlook context (injected
/// as a system message when the popup opens) and the conversation so far.
/// M5 implements this against the Claude API; M4 ships <see cref="PlaceholderDraftGenerator"/>
/// so the popup UI and trigger plumbing can be built and tested independently.
///
/// A null <c>emailContext</c> is not an error: Outlook had nothing open, so the draft is written
/// from what the user says and nothing else.
/// </summary>
public interface IDraftGenerator
{
    /// <param name="sessionId">
    /// Changes every time the popup starts a new conversation. Only <see cref="LocalDraftGenerator"/>
    /// uses it, to tell "the user typed another message" apart from "this is a different
    /// conversation that happens to start the same way" — the two look identical from the arguments
    /// alone once <paramref name="emailContext"/> is allowed to be null, since null equals null.
    /// </param>
    Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        int sessionId = 0,
        CancellationToken cancellationToken = default);
}

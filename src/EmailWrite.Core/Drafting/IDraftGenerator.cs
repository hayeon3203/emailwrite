using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.Core.Drafting;

/// <summary>
/// Produces the assistant's next chat message given the Outlook context (injected
/// as a system message when the popup opens) and the conversation so far.
/// M5 implements this against the Claude API; M4 ships <see cref="PlaceholderDraftGenerator"/>
/// so the popup UI and trigger plumbing can be built and tested independently.
/// </summary>
public interface IDraftGenerator
{
    Task<string> GenerateDraftAsync(
        EmailContext? emailContext,
        IReadOnlyList<ChatMessage> conversation,
        CancellationToken cancellationToken = default);
}

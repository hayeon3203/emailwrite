namespace EmailWrite.Core.OutlookIntegration;

public sealed record EmailMessage(
    string Subject,
    string SenderName,
    string Body,
    DateTime ReceivedTime,
    IReadOnlyList<string> AttachmentNames);

/// <param name="UserName">
/// Who the user is, as Outlook's profile spells it. The prompt lists 발신자 for the mail and for
/// every history entry and never says which of those names is the person writing the reply, so the
/// model picks the most plausible-sounding sentence in the thread and swaps the speaker. Null when
/// Outlook would not tell us, and then the prompt says nothing about it rather than guessing.
/// </param>
public sealed record EmailContext(
    EmailMessage Current,
    IReadOnlyList<EmailMessage> ThreadHistory,
    string? UserName = null);

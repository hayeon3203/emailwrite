namespace EmailWrite.Core.OutlookIntegration;

public sealed record EmailMessage(
    string Subject,
    string SenderName,
    string Body,
    DateTime ReceivedTime,
    IReadOnlyList<string> AttachmentNames);

public sealed record EmailContext(
    EmailMessage Current,
    IReadOnlyList<EmailMessage> ThreadHistory);

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
/// <param name="UserOrg">
/// The user's department, as Exchange spells it — the 소속 half of "다올TS 마케팅팀 조하연입니다".
/// Never reaches the model: it is only used to build the greeting the popup pastes above the draft.
/// Null on a non-Exchange account, or when the user turned down Outlook's access prompt.
/// </param>
public sealed record EmailContext(
    EmailMessage Current,
    IReadOnlyList<EmailMessage> ThreadHistory,
    string? UserName = null,
    string? UserOrg = null);

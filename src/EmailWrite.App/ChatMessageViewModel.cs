using EmailWrite.Core.Drafting;
using EmailWrite.Core.OutlookIntegration;

namespace EmailWrite.App;

/// <summary>Wraps a <see cref="ChatMessage"/> with the bool the bubble DataTemplate binds against.</summary>
public sealed class ChatMessageViewModel(ChatMessage message, bool isLocalOnly = false, EmailContext? greetFor = null)
{
    public ChatMessage Message { get; } = message;

    /// <summary>
    /// What the bubble shows and 복사 puts on the clipboard — the draft with the greeting block on
    /// top when one was asked for. <see cref="Message"/> stays the raw draft, and that is what
    /// <see cref="ChatPopupWindow"/> feeds back as the conversation, so the model never sees the
    /// greeting: it cannot revise it, repeat it, or start treating it as the shape of a reply.
    /// Passing <paramref name="greetFor"/> is what turns this on, and only the draft bubbles do.
    /// </summary>
    public string Text => greetFor == null ? Message.Text : DraftGreeting.Apply(Message.Text, greetFor);
    public bool IsAssistant => Message.Role == ChatRole.Assistant;

    /// <summary>
    /// The popup's own words — which mail it picked up, and errors — rather than anything the
    /// model said. Shown in the thread but kept out of the conversation sent to the model, which
    /// would otherwise be asked to continue from a turn it never produced.
    /// </summary>
    public bool IsLocalOnly { get; } = isLocalOnly;

    /// <summary>
    /// Whether this bubble offers "다시". Only what the model wrote can be rerolled — the popup's
    /// own lines were never generated, and the user's own words are not ours to throw away.
    /// </summary>
    public bool CanRegenerate => IsAssistant && !IsLocalOnly;
}

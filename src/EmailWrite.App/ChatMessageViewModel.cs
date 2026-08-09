using EmailWrite.Core.Drafting;

namespace EmailWrite.App;

/// <summary>Wraps a <see cref="ChatMessage"/> with the bool the bubble DataTemplate binds against.</summary>
public sealed class ChatMessageViewModel(ChatMessage message, bool isLocalOnly = false)
{
    private readonly DraftResponse _parsed = DraftResponse.Parse(message.Text);

    public ChatMessage Message { get; } = message;
    public bool IsAssistant => Message.Role == ChatRole.Assistant;

    /// <summary>
    /// What the bubble shows: the draft without its tags, or the model's own words unchanged.
    /// <see cref="ChatMessage.Text"/> keeps the raw turn, because that is what the model produced
    /// and what it has to see again if the conversation is replayed.
    /// </summary>
    public string Text => _parsed.Text;

    /// <summary>The model called for a draft rather than continuing the conversation.</summary>
    public bool IsDraft => IsAssistant && !IsLocalOnly && _parsed.IsDraft;

    /// <summary>
    /// The popup's own words — which mail it picked up, and errors — rather than anything the
    /// model said. Shown in the thread but kept out of the conversation sent to the model, which
    /// would otherwise be asked to continue from a turn it never produced.
    /// </summary>
    public bool IsLocalOnly { get; } = isLocalOnly;
}

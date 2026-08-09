using EmailWrite.Core.Drafting;

namespace EmailWrite.App;

/// <summary>Wraps a <see cref="ChatMessage"/> with the bool the bubble DataTemplate binds against.</summary>
public sealed class ChatMessageViewModel(ChatMessage message, bool isLocalOnly = false)
{
    public ChatMessage Message { get; } = message;
    public string Text => Message.Text;
    public bool IsAssistant => Message.Role == ChatRole.Assistant;

    /// <summary>
    /// The popup's own words — which mail it picked up, and errors — rather than anything the
    /// model said. Shown in the thread but kept out of the conversation sent to the model, which
    /// would otherwise be asked to continue from a turn it never produced.
    /// </summary>
    public bool IsLocalOnly { get; } = isLocalOnly;
}

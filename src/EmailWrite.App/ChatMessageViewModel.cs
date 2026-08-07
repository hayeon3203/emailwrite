using EmailWrite.Core.Drafting;

namespace EmailWrite.App;

/// <summary>Wraps a <see cref="ChatMessage"/> with the bool the bubble DataTemplate binds against.</summary>
public sealed class ChatMessageViewModel(ChatMessage message)
{
    public ChatMessage Message { get; } = message;
    public string Text => Message.Text;
    public bool IsAssistant => Message.Role == ChatRole.Assistant;
}

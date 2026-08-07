namespace EmailWrite.Core.Drafting;

public enum ChatRole
{
    User,
    Assistant
}

public sealed record ChatMessage(ChatRole Role, string Text, DateTime Timestamp)
{
    public static ChatMessage FromUser(string text) => new(ChatRole.User, text, DateTime.Now);

    public static ChatMessage FromAssistant(string text) => new(ChatRole.Assistant, text, DateTime.Now);
}

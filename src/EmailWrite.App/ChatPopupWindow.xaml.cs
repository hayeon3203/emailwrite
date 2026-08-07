using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using EmailWrite.Core.Drafting;
using EmailWrite.Core.OutlookIntegration;
using Clipboard = System.Windows.Clipboard;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MouseButtonEventArgs = System.Windows.Input.MouseButtonEventArgs;

namespace EmailWrite.App;

/// <summary>
/// Chat popup toggled by the avatar click / global hotkey. Each time it opens,
/// the current Outlook context is captured (<see cref="Open"/>) and handed to
/// <see cref="IDraftGenerator"/> as context for a one-shot draft, alongside the
/// visible conversation.
/// </summary>
public partial class ChatPopupWindow : Window
{
    private readonly IDraftGenerator _draftGenerator;
    private EmailContext? _emailContext;

    public ObservableCollection<ChatMessageViewModel> Messages { get; } = new();

    public ChatPopupWindow() : this(new PlaceholderDraftGenerator())
    {
    }

    public ChatPopupWindow(IDraftGenerator draftGenerator)
    {
        _draftGenerator = draftGenerator;
        InitializeComponent();
        DataContext = this;
    }

    /// <summary>
    /// Resets the conversation, re-reads the current Outlook context, and kicks off one draft.
    /// The Outlook COM round-trip runs off the UI thread — it's a cross-process call that can
    /// take a while, and leaving it on the UI thread freezes the whole app's message loop
    /// (the classic "spinning cursor, nothing happens" hang).
    /// </summary>
    public async Task OpenAsync()
    {
        _emailContext = null;
        Messages.Clear();
        InputBox.Clear();

        TypingIndicator.Visibility = Visibility.Visible;
        ScrollToBottom();

        string? error;
        try
        {
            (_emailContext, error) = await Task.Run(() =>
            {
                OutlookContextReader.TryGetCurrentEmailContext(out var ctx, out var err);
                return (ctx, err);
            });
        }
        catch (Exception ex)
        {
            error = $"Outlook 컨텍스트를 읽는 중 오류가 발생했습니다: {ex.Message}";
        }

        if (_emailContext == null && error != null)
        {
            TypingIndicator.Visibility = Visibility.Collapsed;
            Messages.Add(new ChatMessageViewModel(ChatMessage.FromAssistant(error)));
            ScrollToBottom();
            return;
        }

        await GenerateAndAppendDraftAsync();
    }

    private async Task GenerateAndAppendDraftAsync()
    {
        TypingIndicator.Visibility = Visibility.Visible;
        ScrollToBottom();
        try
        {
            var conversation = Messages.Select(m => m.Message).ToList();
            var draft = await _draftGenerator.GenerateDraftAsync(_emailContext, conversation);
            Messages.Add(new ChatMessageViewModel(ChatMessage.FromAssistant(draft)));
        }
        finally
        {
            TypingIndicator.Visibility = Visibility.Collapsed;
            ScrollToBottom();
        }
    }

    private async void SendButton_Click(object sender, RoutedEventArgs e) => await SendMessageAsync();

    private async void InputBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;

        e.Handled = true;
        await SendMessageAsync();
    }

    private async Task SendMessageAsync()
    {
        var text = InputBox.Text.Trim();
        if (text.Length == 0)
            return;

        InputBox.Clear();
        Messages.Add(new ChatMessageViewModel(ChatMessage.FromUser(text)));
        ScrollToBottom();

        await GenerateAndAppendDraftAsync();
    }

    private void CopyMessage_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { Tag: string text })
            Clipboard.SetText(text);
    }

    private void ScrollToBottom()
    {
        Dispatcher.InvokeAsync(() => MessagesScrollViewer.ScrollToEnd(), DispatcherPriority.Background);
    }

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => DragMove();

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Hide();
}

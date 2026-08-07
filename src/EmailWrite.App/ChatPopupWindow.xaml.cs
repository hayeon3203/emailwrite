using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using EmailWrite.App.Interop;
using EmailWrite.Core.Drafting;
using EmailWrite.Core.OutlookIntegration;
using Button = System.Windows.Controls.Button;
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
        Deactivated += ChatPopupWindow_Deactivated;
    }

    /// <summary>
    /// Auto-hides the popup when it loses focus (e.g. the user clicks elsewhere on
    /// screen), except when the avatar itself was clicked — that click's own
    /// MouseLeftButtonUp handler already owns the show/hide toggle, and hiding here
    /// too would race it into immediately reopening.
    /// </summary>
    private void ChatPopupWindow_Deactivated(object? sender, EventArgs e)
    {
        if (Owner != null && NativeMethods.GetForegroundWindow() == new WindowInteropHelper(Owner).Handle)
            return;

        Hide();
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
            (_emailContext, error) = await ReadOutlookContextAsync();
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

    /// <summary>
    /// Outlook's COM objects are STA-only — calling them from a ThreadPool thread
    /// (which defaults to MTA) makes deep calls like IOleWindow.GetWindow() throw
    /// COMException, which <see cref="OutlookContextReader"/> swallows internally,
    /// silently producing "no mail found" even with a mail open. So this spins up
    /// a dedicated STA thread instead of using Task.Run, keeping the COM round-trip
    /// off the UI thread without breaking apartment affinity.
    /// </summary>
    private static Task<(EmailContext? Context, string? Error)> ReadOutlookContextAsync()
    {
        var tcs = new TaskCompletionSource<(EmailContext?, string?)>();
        var thread = new Thread(() =>
        {
            try
            {
                OutlookContextReader.TryGetCurrentEmailContext(out var ctx, out var err);
                tcs.SetResult((ctx, err));
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        return tcs.Task;
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
        catch (Exception ex)
        {
            // IDraftGenerator implementations are expected to catch their own errors and
            // return a user-facing message string; this is a backstop for anything that
            // still throws (e.g. an unexpected SDK exception) so it surfaces in the chat
            // instead of taking down the popup.
            Messages.Add(new ChatMessageViewModel(ChatMessage.FromAssistant($"초안 생성 중 오류가 발생했어요: {ex.Message}")));
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

    private async void CopyMessage_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string text } button)
            return;

        try
        {
            Clipboard.SetText(text);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Another process is holding the clipboard open (common on Windows); ignore, button just won't flip to "복사됨".
            return;
        }

        var original = button.Content;
        button.Content = "복사됨";
        button.IsEnabled = false;
        await Task.Delay(1000);
        button.Content = original;
        button.IsEnabled = true;
    }

    private void ScrollToBottom()
    {
        Dispatcher.InvokeAsync(() => MessagesScrollViewer.ScrollToEnd(), DispatcherPriority.Background);
    }

    private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => DragMove();

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Hide();
}

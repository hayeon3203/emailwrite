using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
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
    private bool _sessionEnded = true;
    // Bumped on every open so the generator can tell a new conversation from a longer one.
    private int _sessionId;

    public ObservableCollection<ChatMessageViewModel> Messages { get; } = new();

    /// <summary>
    /// True while a draft is being generated. Local inference takes 10+ seconds, long enough that
    /// the popup would otherwise auto-hide mid-generation the moment anything else takes focus —
    /// throwing the draft away and tempting the user into a second F9 that restarts the whole thing.
    /// </summary>
    public bool IsBusy { get; private set; }

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
    /// Whether reopening should read Outlook again and start over, rather than bring back what
    /// was on screen. Hiding keeps the conversation; closing ends it. The popup no longer hides
    /// itself when it loses focus — writing a reply means looking at the mail, at a calendar, at
    /// whatever else, and a window that vanishes on the first click elsewhere cannot be used for
    /// that.
    /// </summary>
    public bool ShouldStartNewSession => !IsBusy && _sessionEnded;

    /// <summary>
    /// Resets the conversation and re-reads the current Outlook context, then waits: the draft is
    /// written from what the user says it should contain, so the popup opens by naming the mail it
    /// picked up and asking for those points rather than guessing a reply from the mail alone.
    /// The Outlook COM round-trip runs off the UI thread — it's a cross-process call that can
    /// take a while, and leaving it on the UI thread freezes the whole app's message loop
    /// (the classic "spinning cursor, nothing happens" hang).
    /// </summary>
    public async Task OpenAsync()
    {
        _emailContext = null;
        Messages.Clear();
        InputBox.Clear();

        _sessionId++;
        _sessionEnded = false;
        IsBusy = true;
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

        IsBusy = false;
        TypingIndicator.Visibility = Visibility.Collapsed;

        // No mail is no longer a dead end. Nothing was open in Outlook, so there is nothing to
        // reply to — but writing the message is still the part worth doing, so the popup says
        // what it found and asks for the content either way. The reason is kept in the line
        // because "Outlook이 실행 중이 아닙니다" and "선택된 메일이 없습니다" are worth telling
        // apart when the user expected a mail to be there.
        if (_emailContext == null)
        {
            // The default hint is about replying to a mail, which is the one thing there isn't
            // one of here.
            InputHint.Text = "예: 내일 오전 회의 30분 늦는다고";
            AddLocalMessage(
                $"{error ?? "열려있거나 선택된 메일이 없습니다."}\n보낼 메시지 내용을 알려주시면 초안을 써드릴게요.");
            return;
        }

        InputHint.Text = "예: 내일까지 검토해서 드린다고";

        // Naming the mail doubles as a check that the right one was picked up — several windows
        // can be open at once, and the reader goes by which was focused last.
        AddLocalMessage(
            $"'{_emailContext.Current.SenderName}'님의 \"{_emailContext.Current.Subject}\" 메일이에요.\n" +
            "답장에 담을 내용을 알려주세요.");
    }

    private void AddLocalMessage(string text)
    {
        Messages.Add(new ChatMessageViewModel(ChatMessage.FromAssistant(text), isLocalOnly: true));
        ScrollToBottom();
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
        IsBusy = true;
        TypingIndicator.Visibility = Visibility.Visible;
        ScrollToBottom();
        try
        {
            var conversation = Messages.Where(m => !m.IsLocalOnly).Select(m => m.Message).ToList();
            var draft = await _draftGenerator.GenerateDraftAsync(_emailContext, conversation, _sessionId);
            Messages.Add(new ChatMessageViewModel(ChatMessage.FromAssistant(draft)));
        }
        catch (Exception ex)
        {
            // IDraftGenerator implementations are expected to catch their own errors and
            // return a user-facing message string; this is a backstop for anything that
            // still throws (e.g. an unexpected SDK exception) so it surfaces in the chat
            // instead of taking down the popup.
            AddLocalMessage($"초안 생성 중 오류가 발생했어요: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
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

    /// <summary>
    /// Throws this draft away and writes another from the same points. Measured on this model, a
    /// quarter of drafts come back unusable while the content behind them was fine — so rerolling
    /// is the ordinary repair, and until now it meant retyping. Restart (<c>↻</c>) is the other
    /// thing: it clears the conversation and re-reads Outlook.
    ///
    /// Everything from this bubble onward goes, so a reroll partway up a long thread also drops
    /// the turns that followed from it — they were answers to a draft that no longer exists.
    /// </summary>
    private async void RegenerateMessage_Click(object sender, RoutedEventArgs e)
    {
        if (IsBusy || sender is not Button { Tag: ChatMessageViewModel clicked })
            return;

        var index = Messages.IndexOf(clicked);
        if (index < 0)
            return;

        while (Messages.Count > index)
        {
            Messages.RemoveAt(Messages.Count - 1);
        }

        // The KV cache still holds the turns just removed, and it can only be appended to. A new
        // session id makes the generator rebuild from the shortened conversation instead.
        _sessionId++;
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

    /// <summary>
    /// Start over without the popup going away. Two things needed this. A conversation can get
    /// stuck — a first message the model answers *about* rather than acts on leaves a turn that
    /// turns every later turn into a refusal — and drafts are unreliable often enough that
    /// throwing one away is a normal move, not an error path. Until now both meant closing the
    /// popup and reopening it. This also re-reads Outlook, so it is the way to move to a
    /// different mail: hiding and reopening keeps the old conversation on purpose.
    /// </summary>
    private async void RestartButton_Click(object sender, RoutedEventArgs e)
    {
        // Mid-generation the reply is still coming and OpenAsync would clear the messages out
        // from under it, leaving the finished draft to land in a conversation it never saw.
        if (IsBusy)
            return;

        await OpenAsync();
    }

    /// <summary>Out of the way, but still there when it comes back.</summary>
    private void HideButton_Click(object sender, RoutedEventArgs e) => Hide();

    /// <summary>Done with this mail: the next open reads Outlook again and starts a new thread.</summary>
    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        _sessionEnded = true;
        Hide();
    }
}

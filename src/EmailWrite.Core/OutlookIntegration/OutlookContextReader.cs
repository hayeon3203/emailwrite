using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailWrite.Core.OutlookIntegration;

// .NET (Core/5+) dropped Marshal.GetActiveObject (it's .NET Framework-only),
// so attaching to an already-running Outlook instance has to go through
// CLSIDFromProgID + the OLE GetActiveObject API directly.
// Not present in System.Runtime.InteropServices.ComTypes on .NET (Core) —
// only in .NET Framework — so declared here. Outlook's Inspector/Explorer
// COM objects implement it, giving access to their underlying HWND.
[ComImport]
[Guid("00000114-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleWindow
{
    void GetWindow(out IntPtr phwnd);
    void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
}

internal static class RunningComObject
{
    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    public static object Get(string progId)
    {
        var hr = CLSIDFromProgID(progId, out var clsid);
        Marshal.ThrowExceptionForHR(hr);
        GetActiveObject(ref clsid, IntPtr.Zero, out var instance);
        return instance;
    }
}

/// <summary>
/// Reads the mail currently open (or selected) in the running classic desktop
/// Outlook via the COM Object Model. Only touches display-safe properties
/// (Subject/SenderName/Body/ReceivedTime/attachment display names) — avoids
/// SenderEmailAddress / Recipients, which can trigger Outlook's "a program is
/// trying to access e-mail information" security prompt.
/// </summary>
public static class OutlookContextReader
{
    private const int MaxThreadHistory = 5;

    public static bool TryGetCurrentEmailContext(out EmailContext? context, out string? error)
    {
        context = null;
        error = null;

        Outlook.Application app;
        try
        {
            app = (Outlook.Application)RunningComObject.Get("Outlook.Application");
        }
        catch (COMException)
        {
            error = "Outlook이 실행 중이 아닙니다.";
            return false;
        }

        var mailItem = GetActiveMailItem(app);
        if (mailItem == null)
        {
            error = "열려있거나 선택된 메일이 없습니다.";
            return false;
        }

        var current = ToEmailMessage(mailItem);
        var history = GetThreadHistory(mailItem);

        context = new EmailContext(current, history);
        return true;
    }

    // If several mail windows (read/compose Inspectors, or Explorer list
    // views with a selection) are open at once, Outlook's own
    // ActiveInspector()/ActiveExplorer() can lag once focus has moved to a
    // non-Outlook window (like this app's popup). So instead: collect every
    // open window that has a mail item in play, and pick whichever one is
    // highest in the real OS Z-order — i.e. was clicked/activated last.
    private static Outlook.MailItem? GetActiveMailItem(Outlook.Application app)
    {
        var candidates = new Dictionary<IntPtr, Outlook.MailItem>();

        foreach (Outlook.Inspector inspector in app.Inspectors)
        {
            if (inspector.CurrentItem is Outlook.MailItem mail)
                TryAddCandidate(candidates, inspector, mail);
        }

        foreach (Outlook.Explorer explorer in app.Explorers)
        {
            if (explorer.Selection is { Count: > 0 } selection && selection[1] is Outlook.MailItem mail)
                TryAddCandidate(candidates, explorer, mail);
        }

        if (candidates.Count == 0)
            return null;

        if (candidates.Count == 1)
            return candidates.Values.First();

        var topmostHwnd = WindowZOrder.FindTopmost(candidates.Keys);
        if (topmostHwnd != IntPtr.Zero && candidates.TryGetValue(topmostHwnd, out var winner))
            return winner;

        // None of the HWND lookups panned out (unexpected) — fall back to
        // Outlook's own idea of "active" rather than returning nothing.
        if (app.ActiveInspector()?.CurrentItem is Outlook.MailItem activeInspectorMail)
            return activeInspectorMail;

        return app.ActiveExplorer()?.Selection is { Count: > 0 } activeSelection &&
               activeSelection[1] is Outlook.MailItem activeSelectedMail
            ? activeSelectedMail
            : null;
    }

    private static void TryAddCandidate(Dictionary<IntPtr, Outlook.MailItem> candidates, object comWindow, Outlook.MailItem mail)
    {
        var hwnd = TryGetHwnd(comWindow);
        if (hwnd != IntPtr.Zero)
            candidates[hwnd] = mail;
    }

    private static IntPtr TryGetHwnd(object comWindow)
    {
        if (comWindow is IOleWindow oleWindow)
        {
            try
            {
                oleWindow.GetWindow(out var hwnd);
                return hwnd;
            }
            catch (COMException)
            {
            }
        }

        return IntPtr.Zero;
    }

    private static EmailMessage ToEmailMessage(Outlook.MailItem mail)
    {
        return new EmailMessage(
            Subject: mail.Subject ?? string.Empty,
            SenderName: mail.SenderName ?? string.Empty,
            Body: mail.Body ?? string.Empty,
            ReceivedTime: mail.ReceivedTime,
            AttachmentNames: GetAttachmentNames(mail));
    }

    // PR_ATTACHMENT_HIDDEN: true for inline images embedded in an HTML
    // signature/body (referenced via cid:) rather than real user attachments.
    private const string PidTagAttachmentHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";

    private static IReadOnlyList<string> GetAttachmentNames(Outlook.MailItem mail)
    {
        var names = new List<string>();
        foreach (Outlook.Attachment attachment in mail.Attachments)
        {
            if (IsHiddenInlineAttachment(attachment))
                continue;

            names.Add(attachment.DisplayName ?? attachment.FileName ?? string.Empty);
        }

        return names;
    }

    private static bool IsHiddenInlineAttachment(Outlook.Attachment attachment)
    {
        try
        {
            return attachment.PropertyAccessor.GetProperty(PidTagAttachmentHidden) is true;
        }
        catch (COMException)
        {
            return false;
        }
    }

    private static IReadOnlyList<EmailMessage> GetThreadHistory(Outlook.MailItem mail)
    {
        var results = new List<EmailMessage>();
        try
        {
            var conversation = mail.GetConversation();
            if (conversation == null)
                return results;

            foreach (object rootItem in conversation.GetRootItems())
                CollectMailItems(conversation, rootItem, results);
        }
        catch (COMException)
        {
            // Conversation tracking isn't available for this store/item — fall back to no history.
        }

        return results
            .OrderByDescending(m => m.ReceivedTime)
            .Take(MaxThreadHistory)
            .ToList();
    }

    private static void CollectMailItems(Outlook.Conversation conversation, object item, List<EmailMessage> results)
    {
        if (item is Outlook.MailItem mail)
            results.Add(ToEmailMessage(mail));

        foreach (object child in conversation.GetChildren(item))
            CollectMailItems(conversation, child, results);
    }
}

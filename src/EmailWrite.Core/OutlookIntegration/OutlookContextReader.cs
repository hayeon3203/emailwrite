using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailWrite.Core.OutlookIntegration;

// .NET (Core/5+) dropped Marshal.GetActiveObject (it's .NET Framework-only),
// so attaching to an already-running Outlook instance has to go through
// CLSIDFromProgID + the OLE GetActiveObject API directly.
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
/// (Subject/SenderName/Body/ReceivedTime) — avoids SenderEmailAddress /
/// Recipients, which can trigger Outlook's "a program is trying to access
/// e-mail information" security prompt.
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

    private static Outlook.MailItem? GetActiveMailItem(Outlook.Application app)
    {
        var inspector = app.ActiveInspector();
        if (inspector?.CurrentItem is Outlook.MailItem inspectorMail)
            return inspectorMail;

        var explorer = app.ActiveExplorer();
        if (explorer?.Selection is { Count: > 0 } selection && selection[1] is Outlook.MailItem selectedMail)
            return selectedMail;

        return null;
    }

    private static EmailMessage ToEmailMessage(Outlook.MailItem mail)
    {
        return new EmailMessage(
            Subject: mail.Subject ?? string.Empty,
            SenderName: mail.SenderName ?? string.Empty,
            Body: mail.Body ?? string.Empty,
            ReceivedTime: mail.ReceivedTime);
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

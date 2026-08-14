using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Threading;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace EmailWrite.App;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    // Held for the process lifetime — a local kept only inside OnStartup gets collected and the
    // guard stops guarding. "Local\" rather than "Global\" is per-session on purpose: two people
    // signed into the same PC each get their own avatar, and a global name would let whoever
    // logged in first lock the second one out.
    private static Mutex? _instanceLock;

    /// <summary>
    /// This is a background tray app with no console — without this handler, an
    /// unhandled exception on the UI thread silently kills the process and the user
    /// just sees the avatar vanish with no explanation.
    /// </summary>
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        // Once this is installed there are several ways to launch it — Start menu, desktop
        // shortcut, and (if it gets added) autostart — so launching a second copy stops being a
        // mistake nobody makes. Two avatars is the visible half; the real damage is that the
        // second process fails to register F9 and pops a warning balloon, so the hotkey looks
        // broken on a PC where it was working. Exits without a message: the first avatar is
        // already on screen, which is the answer to "why did nothing happen".
        _instanceLock = new Mutex(initiallyOwned: true, @"Local\EmailWrite.App", out var isFirstInstance);
        if (!isFirstInstance)
        {
            Shutdown();
        }
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        MessageBox.Show(
            $"예기치 않은 오류가 발생했어요:\n{e.Exception.Message}",
            "EmailWrite",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }
}


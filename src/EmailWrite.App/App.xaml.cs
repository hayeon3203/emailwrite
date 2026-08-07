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
    /// <summary>
    /// This is a background tray app with no console — without this handler, an
    /// unhandled exception on the UI thread silently kills the process and the user
    /// just sees the avatar vanish with no explanation.
    /// </summary>
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnDispatcherUnhandledException;
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


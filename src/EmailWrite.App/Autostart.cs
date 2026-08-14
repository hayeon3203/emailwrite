using Microsoft.Win32;

namespace EmailWrite.App;

/// <summary>
/// Whether Windows launches the avatar at sign-in, stored as a single HKCU Run value.
///
/// Off unless the user turns it on: nothing here writes the value on first run. An app that adds
/// itself to autostart without being asked is what malware does, and antivirus scores it that way.
///
/// The Run key rather than the Startup folder or a scheduled task, because it needs no admin
/// rights and — the part that matters — Windows lists it under 작업 관리자 → 시작 프로그램, so
/// someone who wants it off has a way that does not involve finding this menu. The installer will
/// read and write this same value, which is what keeps an update from resetting the choice.
/// </summary>
internal static class Autostart
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "EmailWrite";

    /// <summary>
    /// Quoted: the installed path is under Program Files, and an unquoted value with a space in it
    /// makes Windows try "C:\Program" first.
    /// </summary>
    private static string? CommandLine =>
        Environment.ProcessPath is { Length: > 0 } path ? $"\"{path}\"" : null;

    internal static bool Enabled
    {
        get
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey);
            return key?.GetValue(ValueName) != null;
        }
        set
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKey);
            if (!value)
            {
                key.DeleteValue(ValueName, throwOnMissingValue: false);
            }
            else if (CommandLine is { } command)
            {
                key.SetValue(ValueName, command, RegistryValueKind.String);
            }
        }
    }

    /// <summary>
    /// Repoints an enabled entry at wherever this copy is actually running from. An update that
    /// lands in a different folder — or a colleague who moves the app — otherwise leaves Windows
    /// launching a path that no longer exists, and the failure is invisible: no error, the avatar
    /// simply stops appearing after a reboot and the menu still says it is on.
    /// Deliberately does not turn autostart on; it only corrects an entry that is already there.
    /// </summary>
    internal static void RepairPath()
    {
        if (CommandLine is not { } command)
            return;

        using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: true);
        if (key?.GetValue(ValueName) is string existing
            && !string.Equals(existing, command, StringComparison.OrdinalIgnoreCase))
        {
            key.SetValue(ValueName, command, RegistryValueKind.String);
        }
    }
}

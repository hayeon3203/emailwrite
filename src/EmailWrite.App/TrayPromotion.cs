using Microsoft.Win32;

namespace EmailWrite.App;

/// <summary>
/// Pins the tray icon to the taskbar instead of leaving it in the "hidden icons" overflow, once,
/// on the first run of this copy.
///
/// Worth doing because the tray icon is the only entry point that survives a reboot in plain
/// sight, and Windows hides new ones by default — a colleague who installs this and right-clicks
/// the avatar away has no visible way back in. Once, and recorded, for the same reason as
/// <see cref="Autostart.EnableOnFirstRun"/>: a value rewritten at every launch makes the Windows
/// setting undoable.
///
/// Undocumented, and Windows 10 keeps this choice somewhere else entirely (an encoded blob under
/// TrayNotify), so this is best-effort by design: when it does not work the icon merely stays in
/// the overflow, which is where it is today.
/// </summary>
internal static class TrayPromotion
{
    private const string SettingsKey = @"Control Panel\NotifyIconSettings";

    /// <summary>
    /// Windows keys each entry by a hash of the executable path that it computes itself, and only
    /// writes the entry after the icon has actually appeared — so the installer cannot create it
    /// ahead of time, and neither can this before the icon is shown. Measured at 1.1s after
    /// launch on this machine; polled rather than read once so a slower PC is not a silent miss.
    /// </summary>
    internal static async Task PromoteOnFirstRunAsync()
    {
        try
        {
            using var marker = Registry.CurrentUser.CreateSubKey(@"Software\EmailWrite");
            if (marker.GetValue("TrayPromoted") != null || Environment.ProcessPath is not { Length: > 0 } exe)
            {
                return;
            }

            for (var attempt = 0; attempt < 20; attempt++)
            {
                using (var all = Registry.CurrentUser.OpenSubKey(SettingsKey))
                {
                    foreach (var name in all?.GetSubKeyNames() ?? [])
                    {
                        using var entry = all!.OpenSubKey(name, writable: true);
                        if (entry?.GetValue("ExecutablePath") is string path
                            && string.Equals(path, exe, StringComparison.OrdinalIgnoreCase))
                        {
                            entry.SetValue("IsPromoted", 1, RegistryValueKind.DWord);
                            // Only after it actually landed: giving up early would burn the one
                            // attempt on a PC where the entry was merely slow to appear.
                            marker.SetValue("TrayPromoted", 1, RegistryValueKind.DWord);
                            return;
                        }
                    }
                }

                await Task.Delay(250);
            }
        }
        catch (Exception)
        {
            // Cosmetic, and reached from a fire-and-forget task: a locked-down or restructured
            // registry must not take the avatar down with it.
        }
    }
}

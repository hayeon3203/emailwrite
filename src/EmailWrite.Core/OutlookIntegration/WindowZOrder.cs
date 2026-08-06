using System.Runtime.InteropServices;

namespace EmailWrite.Core.OutlookIntegration;

/// <summary>
/// Resolves which of several candidate windows was activated most recently,
/// using the real OS Z-order rather than trusting any single app's internal
/// "active window" bookkeeping (which can lag once focus moves outside it).
/// </summary>
internal static class WindowZOrder
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    /// <summary>
    /// EnumWindows visits top-level windows top-to-bottom in Z-order, so the
    /// first candidate it hits is the one highest in the stack — i.e. the
    /// one most recently brought to front/clicked among the candidates.
    /// Returns IntPtr.Zero if none of the candidates are live top-level windows.
    /// </summary>
    public static IntPtr FindTopmost(IReadOnlyCollection<IntPtr> candidates)
    {
        if (candidates.Count == 0)
            return IntPtr.Zero;

        var found = IntPtr.Zero;
        EnumWindows((hWnd, _) =>
        {
            if (!candidates.Contains(hWnd))
                return true;

            found = hWnd;
            return false;
        }, IntPtr.Zero);

        return found;
    }
}

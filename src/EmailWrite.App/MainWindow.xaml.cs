using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using EmailWrite.App.Interop;
using Application = System.Windows.Application;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Point = System.Windows.Point;

namespace EmailWrite.App;

/// <summary>
/// Desktop avatar overlay: idle-animated, draggable, always-on-top.
/// Click or the F9 global hotkey toggles the chat popup.
/// </summary>
public partial class MainWindow : Window
{
    private const double EdgeMargin = 24;
    private static readonly TimeSpan FrameInterval = TimeSpan.FromMilliseconds(150);

    private const int HotkeyId = 0x0900;
    private const uint VK_F9 = 0x78;

    private readonly string _framesDir = Path.Combine(AppContext.BaseDirectory, "Assets", "Avatar", "Frames");

    private BitmapImage[] _frames = [];
    private int _frameIndex;
    private DispatcherTimer? _animationTimer;
    private System.Windows.Forms.NotifyIcon? _trayIcon;
    private Point _mouseDownPosition;
    private bool _isDragging;
    private ChatPopupWindow? _chatPopup;

    public event EventHandler? AvatarClicked;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Closed += OnClosed;
        AvatarClicked += (_, _) => ToggleChatPopup();
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        PositionAtBottomRight();
        LoadFrames();
        StartIdleAnimation();
        SetupTrayIcon();
        RegisterGlobalHotkey();
        Autostart.RepairPath();
    }

    private void RegisterGlobalHotkey()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        HwndSource.FromHwnd(hwnd)?.AddHook(WndProc);

        if (!NativeMethods.RegisterHotKey(hwnd, HotkeyId, 0, VK_F9))
        {
            _trayIcon?.ShowBalloonTip(
                3000,
                "EmailWrite",
                "F9 단축키가 다른 프로그램에서 이미 사용 중이라 등록하지 못했습니다.",
                System.Windows.Forms.ToolTipIcon.Warning);
        }
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == NativeMethods.WM_HOTKEY && wParam.ToInt32() == HotkeyId)
        {
            ToggleChatPopup();
            handled = true;
        }

        return IntPtr.Zero;
    }

    private void ToggleChatPopup()
    {
        if (_chatPopup is { IsVisible: true })
        {
            _chatPopup.Hide();
            return;
        }

        // The backend is picked in the popup's own header dropdown now (EMAILWRITE_BACKEND still
        // decides what it opens on).
        _chatPopup ??= new ChatPopupWindow { Owner = this };
        PositionChatPopupAboveAvatar();
        _chatPopup.Show();
        _chatPopup.Activate();

        // Bringing a hidden popup back should show what was there, not throw it away: only a
        // closed (or never opened) one starts over, and one still generating just comes back
        // into view.
        if (_chatPopup.ShouldStartNewSession)
            _ = _chatPopup.OpenAsync();
    }

    /// <summary>
    /// Centred over the avatar, then pulled back inside the work area. Centring alone put the
    /// popup 76px off the right edge in the default position — the avatar sits 24px from the
    /// right, and a 320-wide popup centred on a 120-wide avatar reaches 100px further right than
    /// the avatar does. The same applies upward once the avatar has been dragged near the top.
    /// </summary>
    private void PositionChatPopupAboveAvatar()
    {
        if (_chatPopup == null)
            return;

        const double gap = 12;
        var workArea = SystemParameters.WorkArea;
        _chatPopup.Left = KeepInside(
            Left + (Width / 2) - (_chatPopup.Width / 2), _chatPopup.Width, workArea.Left, workArea.Right);
        _chatPopup.Top = KeepInside(
            Top - _chatPopup.Height - gap, _chatPopup.Height, workArea.Top, workArea.Bottom);
    }

    /// <summary>
    /// Nudges a window edge back between <paramref name="min"/> and <paramref name="max"/>.
    /// Not Math.Clamp: that throws when the window is larger than the space it has to fit in,
    /// which is a screen small enough to be someone's problem but not a reason to crash.
    /// </summary>
    private static double KeepInside(double start, double length, double min, double max) =>
        Math.Max(min, Math.Min(start, max - length));

    private void PositionAtBottomRight()
    {
        var workArea = SystemParameters.WorkArea;
        Left = workArea.Right - Width - EdgeMargin;
        Top = workArea.Bottom - Height - EdgeMargin;
    }

    private void LoadFrames()
    {
        if (!Directory.Exists(_framesDir))
            return;

        var files = Directory.GetFiles(_framesDir, "frame_*.png");
        Array.Sort(files, StringComparer.Ordinal);
        _frames = Array.ConvertAll(files, f => new BitmapImage(new Uri(f, UriKind.Absolute)));

        if (_frames.Length > 0)
            AvatarImage.Source = _frames[0];
    }

    private void StartIdleAnimation()
    {
        if (_frames.Length == 0)
            return;

        _animationTimer = new DispatcherTimer { Interval = FrameInterval };
        _animationTimer.Tick += (_, _) =>
        {
            _frameIndex = (_frameIndex + 1) % _frames.Length;
            AvatarImage.Source = _frames[_frameIndex];
        };
        _animationTimer.Start();
    }

    private void SetupTrayIcon()
    {
        var iconPath = Path.Combine(_framesDir, "frame_00.png");
        System.Drawing.Icon trayIcon;
        if (File.Exists(iconPath))
        {
            using var bitmap = new Bitmap(iconPath);
            trayIcon = System.Drawing.Icon.FromHandle(bitmap.GetHicon());
        }
        else
        {
            trayIcon = SystemIcons.Application;
        }

        var contextMenu = new System.Windows.Forms.ContextMenuStrip();

        // Read on every open rather than once here: this value is also reachable from regedit,
        // from the installer, and from a second copy of this menu, so a cached checkmark would
        // eventually disagree with what Windows will actually do.
        var autostartItem = new System.Windows.Forms.ToolStripMenuItem("Windows 시작 시 실행")
        {
            CheckOnClick = true
        };
        autostartItem.Click += (_, _) => Autostart.Enabled = autostartItem.Checked;
        contextMenu.Opening += (_, _) => autostartItem.Checked = Autostart.Enabled;
        contextMenu.Items.Add(autostartItem);

        contextMenu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
        contextMenu.Items.Add("종료", null, (_, _) => Application.Current.Shutdown());

        _trayIcon = new System.Windows.Forms.NotifyIcon
        {
            Icon = trayIcon,
            Visible = true,
            Text = "EmailWrite Avatar",
            ContextMenuStrip = contextMenu
        };
    }

    // MouseLeftButtonDown/MouseMove/MouseLeftButtonUp together distinguish a
    // click (no movement) from a drag, since DragMove() would otherwise
    // swallow the click that M2 needs to open the chat popup.
    private void AvatarImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _mouseDownPosition = e.GetPosition(this);
        _isDragging = false;
        AvatarImage.CaptureMouse();
    }

    private void AvatarImage_MouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isDragging)
            return;

        var delta = e.GetPosition(this) - _mouseDownPosition;
        if (Math.Abs(delta.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(delta.Y) < SystemParameters.MinimumVerticalDragDistance)
            return;

        _isDragging = true;
        AvatarImage.ReleaseMouseCapture();
        DragMove();
    }

    private void AvatarImage_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        AvatarImage.ReleaseMouseCapture();
        if (!_isDragging)
            AvatarClicked?.Invoke(this, EventArgs.Empty);
    }

    // Right-click on the avatar quits, same as the tray menu's 종료. The tray icon is the only
    // other way out and it is easy to lose track of; the avatar is the thing on screen.
    private void AvatarImage_MouseRightButtonUp(object sender, MouseButtonEventArgs e) =>
        Application.Current.Shutdown();

    private void OnClosed(object? sender, EventArgs e)
    {
        NativeMethods.UnregisterHotKey(new WindowInteropHelper(this).Handle, HotkeyId);
        _animationTimer?.Stop();
        if (_trayIcon != null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
        }
    }
}

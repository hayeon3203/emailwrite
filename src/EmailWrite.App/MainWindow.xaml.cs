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

        _chatPopup ??= new ChatPopupWindow(new EmailWrite.Core.Drafting.LocalDraftGenerator()) { Owner = this };
        PositionChatPopupAboveAvatar();
        _chatPopup.Show();
        _chatPopup.Activate();

        // Re-showing a popup that's still generating just brings it back into view — restarting
        // OpenAsync here would throw away the in-flight draft and queue a second inference behind it.
        if (!_chatPopup.IsBusy)
            _ = _chatPopup.OpenAsync();
    }

    private void PositionChatPopupAboveAvatar()
    {
        if (_chatPopup == null)
            return;

        const double gap = 12;
        _chatPopup.Left = Left + (Width / 2) - (_chatPopup.Width / 2);
        _chatPopup.Top = Top - _chatPopup.Height - gap;
    }

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

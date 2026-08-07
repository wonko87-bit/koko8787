using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Flowdeck.Core.Settings;

namespace Flowdeck.Windows.Interop;

/// <summary>
/// Keeps the widget where the user asked for it in the window stack.
///
/// <see cref="WidgetPinMode.Desktop"/> is the interesting one: the window is pushed to
/// the bottom of the z-order whenever anything else is in front, so it reads as part of
/// the desktop rather than as a floating window — but it is still allowed to come forward
/// while it holds focus, which is what makes its inline input box usable.
/// It is also marked as a tool window so it stays out of Alt+Tab and the taskbar.
/// </summary>
public sealed class WindowPinService : IDisposable
{
    private readonly Window _window;
    private HwndSource? _source;
    private WidgetPinMode _mode = WidgetPinMode.Desktop;

    public WindowPinService(Window window)
    {
        _window = window;
        _window.SourceInitialized += OnSourceInitialized;
    }

    public void Apply(WidgetPinMode mode)
    {
        _mode = mode;
        if (_source is null) return;

        _window.Topmost = mode == WidgetPinMode.AlwaysOnTop;

        if (mode == WidgetPinMode.Desktop) SendToBottom();
    }

    private void OnSourceInitialized(object? sender, EventArgs e)
    {
        _source = (HwndSource)PresentationSource.FromVisual(_window)!;
        _source.AddHook(WndProc);

        HideFromAltTab(_source.Handle);
        Apply(_mode);
    }

    /// <summary>Adds WS_EX_TOOLWINDOW, which takes the window out of Alt+Tab and the taskbar.</summary>
    private static void HideFromAltTab(IntPtr handle)
    {
        var style = NativeMethods.GetWindowLong(handle, NativeMethods.GWL_EXSTYLE);
        style |= NativeMethods.WS_EX_TOOLWINDOW;
        style &= ~NativeMethods.WS_EX_APPWINDOW;
        NativeMethods.SetWindowLong(handle, NativeMethods.GWL_EXSTYLE, style);
    }

    private void SendToBottom()
    {
        if (_source is null) return;

        NativeMethods.SetWindowPos(
            _source.Handle,
            NativeMethods.HWND_BOTTOM,
            0, 0, 0, 0,
            NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (_mode != WidgetPinMode.Desktop || msg != NativeMethods.WM_WINDOWPOSCHANGING) return IntPtr.Zero;

        // While the widget is the foreground window the user is interacting with it,
        // so leave the requested z-order alone. Otherwise force it back to the bottom.
        if (NativeMethods.GetForegroundWindow() == hwnd) return IntPtr.Zero;

        var position = Marshal.PtrToStructure<NativeMethods.WINDOWPOS>(lParam);
        position.hwndInsertAfter = NativeMethods.HWND_BOTTOM;
        position.flags &= ~NativeMethods.SWP_NOZORDER;
        Marshal.StructureToPtr(position, lParam, fDeleteOld: false);

        return IntPtr.Zero;
    }

    public void Dispose()
    {
        _window.SourceInitialized -= OnSourceInitialized;
        _source?.RemoveHook(WndProc);
        _source = null;
    }
}

using System.Windows.Input;
using System.Windows.Interop;

namespace Flowdeck.Windows.Interop;

/// <summary>Why a hot key could not be claimed.</summary>
public enum HotKeyResult
{
    Registered,
    InvalidGesture,

    /// <summary>Another application already owns the combination.</summary>
    AlreadyTaken,
}

/// <summary>
/// System-wide hot keys.
///
/// The keys are attached to a message-only window rather than to any of the app's
/// visible windows, so they keep working while the widget is hidden and survive the
/// widget being closed and reopened.
/// </summary>
public sealed class HotKeyService : IDisposable
{
    private readonly Dictionary<int, Action> _handlers = new();
    private readonly HwndSource _source;
    private int _nextId = 1;
    private bool _disposed;

    public HotKeyService()
    {
        var parameters = new HwndSourceParameters("Flowdeck.HotKeys")
        {
            Width = 0,
            Height = 0,
            ParentWindow = NativeMethods.HWND_MESSAGE,
            WindowStyle = 0,
        };

        _source = new HwndSource(parameters);
        _source.AddHook(WndProc);
    }

    /// <summary>
    /// Claims <paramref name="gesture"/>, e.g. "Ctrl+Alt+Space". Existing registrations
    /// are left alone; call <see cref="UnregisterAll"/> first when re-applying settings.
    /// </summary>
    public HotKeyResult Register(string gesture, Action handler)
    {
        if (!TryParse(gesture, out var modifiers, out var virtualKey)) return HotKeyResult.InvalidGesture;

        var id = _nextId++;
        if (!NativeMethods.RegisterHotKey(_source.Handle, id, modifiers | NativeMethods.MOD_NOREPEAT, virtualKey))
        {
            return HotKeyResult.AlreadyTaken;
        }

        _handlers[id] = handler;
        return HotKeyResult.Registered;
    }

    public void UnregisterAll()
    {
        foreach (var id in _handlers.Keys)
        {
            NativeMethods.UnregisterHotKey(_source.Handle, id);
        }

        _handlers.Clear();
    }

    /// <summary>
    /// Reads a gesture string into Win32 modifier flags and a virtual key code.
    /// Accepts "Ctrl", "Control", "Alt", "Shift", "Win" in any order and any case.
    /// </summary>
    public static bool TryParse(string? gesture, out uint modifiers, out uint virtualKey)
    {
        modifiers = 0;
        virtualKey = 0;
        if (string.IsNullOrWhiteSpace(gesture)) return false;

        var parts = gesture.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0) return false;

        Key? key = null;

        foreach (var part in parts)
        {
            switch (part.ToLowerInvariant())
            {
                case "ctrl":
                case "control":
                    modifiers |= NativeMethods.MOD_CONTROL;
                    break;

                case "alt":
                    modifiers |= NativeMethods.MOD_ALT;
                    break;

                case "shift":
                    modifiers |= NativeMethods.MOD_SHIFT;
                    break;

                case "win":
                case "windows":
                    modifiers |= NativeMethods.MOD_WIN;
                    break;

                default:
                    if (key is not null) return false; // two non-modifier keys
                    if (!Enum.TryParse<Key>(part, ignoreCase: true, out var parsed)) return false;
                    key = parsed;
                    break;
            }
        }

        // A bare key with no modifier would swallow that key for the whole desktop.
        if (key is null || modifiers == 0) return false;

        virtualKey = (uint)KeyInterop.VirtualKeyFromKey(key.Value);
        return virtualKey != 0;
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != NativeMethods.WM_HOTKEY) return IntPtr.Zero;

        if (_handlers.TryGetValue(wParam.ToInt32(), out var handler))
        {
            handled = true;
            handler();
        }

        return IntPtr.Zero;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        UnregisterAll();
        _source.RemoveHook(WndProc);
        _source.Dispose();
    }
}

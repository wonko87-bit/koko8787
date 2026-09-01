using System.Runtime.InteropServices;
using System.Windows;

namespace Flowdeck.Windows.Interop;

/// <summary>
/// Reads whatever text is selected in the window in front.
///
/// Windows has no way to ask an application what it has selected — the right-click menu over
/// a selection is drawn by that application, and nothing of the system's reaches inside it.
/// What every application does agree on is Ctrl+C, so that is what this sends, and then it
/// reads the clipboard. Crude, and it works in Outlook, a browser, a PDF reader and Notepad
/// alike, which no per-application integration would.
///
/// The clipboard keeps what was copied. Restoring what was there before would mean putting
/// back the plain text of something that may have been a picture or a formatted block, and
/// quietly replacing the one with the other is worse than leaving the copy in place.
/// </summary>
public static class SelectionReader
{
    private const uint InputKeyboard = 1;
    private const uint KeyUp = 0x0002;

    private const ushort VkShift = 0x10;
    private const ushort VkControl = 0x11;
    private const ushort VkAlt = 0x12;
    private const ushort VkLeftWindows = 0x5B;
    private const ushort VkRightWindows = 0x5C;
    private const ushort VkC = 0x43;

    /// <summary>How long to give the other application to answer, and how often to look.</summary>
    private static readonly TimeSpan Patience = TimeSpan.FromMilliseconds(400);
    private static readonly TimeSpan Beat = TimeSpan.FromMilliseconds(20);

    /// <summary>
    /// The selected text, or what is already on the clipboard when nothing could be read —
    /// which covers the application that does not copy on Ctrl+C, and the person who copied
    /// it themselves a moment ago. Null when there is no text to be had either way.
    /// </summary>
    public static async Task<string?> ReadAsync()
    {
        var before = GetClipboardSequenceNumber();

        SendCopy();

        // Polling the sequence number rather than sleeping a fixed spell: a local text box
        // answers at once and a mail client under load takes its time, and this is quick for
        // the first without giving up on the second.
        for (var waited = TimeSpan.Zero; waited < Patience; waited += Beat)
        {
            await Task.Delay(Beat);
            if (GetClipboardSequenceNumber() != before) break;
        }

        return Text();
    }

    /// <summary>
    /// Ctrl+C, with the keys that summoned us let go of first.
    ///
    /// The hot key is a chord, so its modifiers are still physically held when this runs. Left
    /// as they are, the application in front would see Ctrl+Alt+C rather than Ctrl+C and do
    /// something else entirely, or nothing.
    /// </summary>
    private static void SendCopy()
    {
        var keys = new List<Input>
        {
            Up(VkAlt),
            Up(VkShift),
            Up(VkLeftWindows),
            Up(VkRightWindows),
            Up(VkControl),
            Down(VkControl),
            Down(VkC),
            Up(VkC),
            Up(VkControl),
        };

        var batch = keys.ToArray();
        SendInput((uint)batch.Length, batch, Marshal.SizeOf<Input>());
    }

    /// <summary>
    /// The clipboard belongs to whoever has it open, so a read can simply fail while another
    /// application is mid-write. Worth a few attempts and not worth an exception.
    /// </summary>
    private static string? Text()
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                if (!Clipboard.ContainsText()) return null;

                var text = Clipboard.GetText();
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }
            catch (Exception e) when (e is ExternalException or InvalidOperationException)
            {
                Thread.Sleep(20);
            }
        }

        return null;
    }

    private static Input Down(ushort key) => new()
    {
        Type = InputKeyboard,
        Data = new InputData { Keyboard = new KeyboardInput { VirtualKey = key } },
    };

    private static Input Up(ushort key) => new()
    {
        Type = InputKeyboard,
        Data = new InputData { Keyboard = new KeyboardInput { VirtualKey = key, Flags = KeyUp } },
    };

    // ---- interop -----------------------------------------------------------

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint count, Input[] inputs, int size);

    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();

    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public uint Type;
        public InputData Data;
    }

    /// <summary>
    /// The three shapes SendInput accepts, overlaid. All three are declared even though only
    /// the keyboard one is ever filled in: the union is as wide as its widest member, and a
    /// struct that left the others out would be the wrong size for the call.
    /// </summary>
    [StructLayout(LayoutKind.Explicit)]
    private struct InputData
    {
        [FieldOffset(0)] public MouseInput Mouse;
        [FieldOffset(0)] public KeyboardInput Keyboard;
        [FieldOffset(0)] public HardwareInput Hardware;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int X;
        public int Y;
        public uint Data;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput
    {
        public ushort VirtualKey;
        public ushort ScanCode;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HardwareInput
    {
        public uint Message;
        public ushort LowOrder;
        public ushort HighOrder;
    }
}

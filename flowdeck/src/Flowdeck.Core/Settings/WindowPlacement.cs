namespace Flowdeck.Core.Settings;

/// <summary>
/// Where a window last sat. <see cref="Left"/> and <see cref="Top"/> are NaN until the
/// window has been positioned once, which is the signal to centre it instead.
/// </summary>
public sealed class WindowPlacement
{
    public double Left { get; set; } = double.NaN;

    public double Top { get; set; } = double.NaN;

    public double Width { get; set; } = 420;

    public double Height { get; set; } = 620;

    public bool HasPosition => !double.IsNaN(Left) && !double.IsNaN(Top);
}

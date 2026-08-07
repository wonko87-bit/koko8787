using System.Text.Json.Serialization;

namespace Flowdeck.Core.Settings;

/// <summary>
/// Where a window last sat. <see cref="Left"/> and <see cref="Top"/> are null until the
/// window has been positioned once, which is the signal to centre it instead.
///
/// Null rather than NaN: System.Text.Json refuses to write NaN, so a sentinel of that
/// kind takes the whole settings file down with it the first time anything is saved.
/// </summary>
public sealed class WindowPlacement
{
    public double? Left { get; set; }

    public double? Top { get; set; }

    public double Width { get; set; } = 420;

    public double Height { get; set; } = 620;

    [JsonIgnore]
    public bool HasPosition => Left.HasValue && Top.HasValue;
}

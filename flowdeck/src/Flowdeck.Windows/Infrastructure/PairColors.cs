using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Flowdeck.Windows.Infrastructure;

/// <summary>
/// The colours a linked event and its todos share on screen.
///
/// Ten, hand-picked rather than computed. Spacing hues evenly round the wheel looks
/// principled and is not: a third of the wheel is yellow-green to the eye, and the "evenly
/// spaced" set lands three near-identical colours there. This set was chosen for the
/// distance between every pair as seen, holds up against the common red-green weakness
/// reasonably, and sits at a middle lightness that reads on the dark theme and the light
/// one alike. The order alternates families so that neighbouring indices never resemble
/// each other — the first two pairs on a screen are blue and orange, not blue and teal.
///
/// A colour is assigned to a pair for one screen at a time, not stored with it. Colour is
/// not identity; it is how two rows find each other in the same list, and a list never
/// shows enough pairs for ten to run out.
/// </summary>
public static class PairColors
{
    private static readonly Color[] Colors =
    {
        Color.FromRgb(0x4E, 0x79, 0xA7), // blue
        Color.FromRgb(0xF2, 0x8E, 0x2B), // orange
        Color.FromRgb(0x59, 0xA1, 0x4F), // green
        Color.FromRgb(0xE1, 0x57, 0x59), // red
        Color.FromRgb(0xB0, 0x7A, 0xA1), // purple
        Color.FromRgb(0x17, 0xBE, 0xCF), // cyan
        Color.FromRgb(0xED, 0xC9, 0x48), // yellow
        Color.FromRgb(0x9C, 0x75, 0x5F), // brown
        Color.FromRgb(0xFF, 0x9D, 0xA7), // pink
        Color.FromRgb(0x76, 0xB7, 0xB2), // teal
    };

    private static readonly Brush[] Brushes = Colors
        .Select(c => { var b = new SolidColorBrush(c); b.Freeze(); return (Brush)b; })
        .ToArray();

    public static int Count => Colors.Length;

    /// <summary>The brush for a pair index, or nothing for a row that is in no pair.</summary>
    public static Brush For(int index) =>
        index < 0 ? System.Windows.Media.Brushes.Transparent : Brushes[index % Brushes.Length];
}

/// <summary>Turns a row's pair index into the brush for its bar. -1 is no bar, as an invisible one.</summary>
public sealed class PairIndexToBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        PairColors.For(value is int index ? index : -1);

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

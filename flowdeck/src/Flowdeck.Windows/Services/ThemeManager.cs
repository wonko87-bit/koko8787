using System.Windows;
using Flowdeck.Core.Settings;

namespace Flowdeck.Windows.Services;

/// <summary>
/// Swaps the palette dictionary in place. Every control binds its colours with
/// DynamicResource, so the whole app re-themes without reloading a window.
/// </summary>
public static class ThemeManager
{
    /// <summary>The palette always occupies the first merged dictionary slot.</summary>
    private const int PaletteSlot = 0;

    public static void Apply(AppTheme theme)
    {
        var name = theme == AppTheme.Light ? "Light" : "Dark";
        var source = new Uri($"pack://application:,,,/Themes/{name}.xaml", UriKind.Absolute);

        var dictionaries = Application.Current.Resources.MergedDictionaries;
        var palette = new ResourceDictionary { Source = source };

        if (dictionaries.Count > PaletteSlot)
        {
            dictionaries[PaletteSlot] = palette;
        }
        else
        {
            dictionaries.Insert(PaletteSlot, palette);
        }
    }
}

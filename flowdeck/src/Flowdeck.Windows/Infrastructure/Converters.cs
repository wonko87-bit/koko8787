using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Flowdeck.Windows.Infrastructure;

/// <summary>
/// Collapses on false. Pass "Invert" as the parameter to flip it, and "Hidden"
/// to reserve the layout space instead of removing it.
/// </summary>
public sealed class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var flag = value is bool b && b;
        var option = parameter as string ?? string.Empty;

        if (option.Contains("Invert", StringComparison.OrdinalIgnoreCase)) flag = !flag;

        var hiddenState = option.Contains("Hidden", StringComparison.OrdinalIgnoreCase)
            ? Visibility.Hidden
            : Visibility.Collapsed;

        return flag ? Visibility.Visible : hiddenState;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is Visibility visibility && visibility == Visibility.Visible;
}

/// <summary>
/// Collapses when the bound string is null, empty or whitespace. Pass "Invert" to
/// show only while the string is empty, which is how the placeholder text works.
/// </summary>
public sealed class StringToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var hasText = !string.IsNullOrWhiteSpace(value as string);

        if (parameter is string option && option.Contains("Invert", StringComparison.OrdinalIgnoreCase))
        {
            hasText = !hasText;
        }

        return hasText ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

using System.Text.Json;
using System.Text.Json.Serialization;
using Flowdeck.Core.Parsing;

namespace Flowdeck.Core.Settings;

public enum AppTheme
{
    Dark = 0,
    Light = 1,
}

/// <summary>How the widget sits relative to other windows.</summary>
public enum WidgetPinMode
{
    /// <summary>Sits on the desktop, below ordinary windows. The default.</summary>
    Desktop = 0,

    /// <summary>Floats above everything.</summary>
    AlwaysOnTop = 1,

    /// <summary>Behaves like a normal window.</summary>
    Normal = 2,
}

/// <summary>User preferences. Serialised next to the workspace file.</summary>
public sealed class AppSettings
{
    public AppTheme Theme { get; set; } = AppTheme.Dark;

    /// <summary>Opens the natural-language input overlay. Parsed by the Windows layer.</summary>
    public string QuickAddHotkey { get; set; } = "Ctrl+Alt+Space";

    /// <summary>Shows or hides the desktop widget.</summary>
    public string ToggleWidgetHotkey { get; set; } = "Ctrl+Alt+D";

    /// <summary>Opens or closes the full calendar list. C for calendar.</summary>
    public string AgendaEventsHotkey { get; set; } = "Ctrl+Alt+C";

    /// <summary>Opens or closes the full todo list.</summary>
    public string AgendaTodosHotkey { get; set; } = "Ctrl+Alt+T";

    /// <summary>Whether the todo list window includes items that are already done.</summary>
    public bool AgendaShowCompleted { get; set; }

    /// <summary>
    /// With several tags picked in a list window, true narrows to entries carrying all of
    /// them and false widens to entries carrying any.
    /// </summary>
    public bool AgendaMatchAllTags { get; set; } = true;

    public WindowPlacement AgendaEventsPlacement { get; set; } = new();

    public WindowPlacement AgendaTodosPlacement { get; set; } = new();

    public WidgetPinMode PinMode { get; set; } = WidgetPinMode.Desktop;

    public double WidgetLeft { get; set; } = double.NaN;

    public double WidgetTop { get; set; } = double.NaN;

    public double WidgetWidth { get; set; } = 340;

    /// <summary>Derived from the width by the widget's fixed aspect ratio; stored for reference.</summary>
    public double WidgetHeight { get; set; } = 551;

    /// <summary>0.35 – 1.0. Applied to the widget's background, never to its text.</summary>
    public double WidgetOpacity { get; set; } = 0.92;

    public bool ShowWidgetOnStart { get; set; } = true;

    public bool LaunchAtStartup { get; set; }

    /// <summary>Hide undated todos in the widget list.</summary>
    public bool HideUndatedTodos { get; set; }

    /// <summary>
    /// Which column the month grid starts on. This also decides what "이번주 금요일" means,
    /// so the parser and the calendar always agree about where a week begins.
    /// </summary>
    public DayOfWeek FirstDayOfWeek { get; set; } = DayOfWeek.Monday;

    /// <summary>
    /// Opt in to a blue Saturday and a red Sunday. Off by default: the palette is
    /// otherwise monochrome, and weekends are already set apart by a lighter grey.
    /// </summary>
    public bool ColorWeekends { get; set; }

    /// <summary>
    /// Reads a bare "2시" as 14:00 rather than 02:00. See
    /// <c>NaturalLanguageParser.AssumeAfternoonForBareHours</c>.
    /// </summary>
    public bool AssumeAfternoonForBareHours { get; set; } = true;

    /// <summary>Editable so the routing vocabulary can be tuned without a rebuild.</summary>
    public RoutingRules Routing { get; set; } = new();

    [JsonIgnore]
    public static string DefaultFilePath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Flowdeck",
            "settings.json");

    public AppSettings Clone() => Load(JsonSerializer.Serialize(this, SerializerOptions)) ?? new AppSettings();

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    public static AppSettings? Load(string json) =>
        JsonSerializer.Deserialize<AppSettings>(json, SerializerOptions);

    public static AppSettings LoadFrom(string path)
    {
        try
        {
            if (!File.Exists(path)) return new AppSettings();
            return Load(File.ReadAllText(path)) ?? new AppSettings();
        }
        catch (Exception e) when (e is IOException or JsonException or UnauthorizedAccessException)
        {
            return new AppSettings();
        }
    }

    public void SaveTo(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllText(path, JsonSerializer.Serialize(this, SerializerOptions));
    }
}

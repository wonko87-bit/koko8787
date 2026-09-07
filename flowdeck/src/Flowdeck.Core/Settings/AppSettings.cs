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

/// <summary>Which meetings get a "want to write a line?" prompt when they end.</summary>
public enum MeetingNudgeScope
{
    /// <summary>Every meeting on the Outlook overlay, and every copy taken. The default.</summary>
    All = 0,

    /// <summary>Only meetings the user has already taken a copy of.</summary>
    Adopted = 1,

    Off = 2,
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

    /// <summary>
    /// Opens the input overlay filled with whatever text is selected in the window in front.
    /// V because it is the paste key, and this arrives at the same place a paste would.
    /// </summary>
    public string CaptureSelectionHotkey { get; set; } = "Ctrl+Alt+V";

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

    /// <summary>Null until the widget has been placed once. Never NaN — see <see cref="WindowPlacement"/>.</summary>
    public double? WidgetLeft { get; set; }

    public double? WidgetTop { get; set; }

    public double WidgetWidth { get; set; } = 340;

    /// <summary>Independent of the width; every pixel above the minimum goes to the list.</summary>
    public double WidgetHeight { get; set; } = 620;

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

    /// <summary>
    /// Master switch for the Outlook integration. Turning it off makes Flowdeck purely
    /// local, as it was before the integration existed. Only ever consulted on a machine
    /// that has Outlook: where it is absent the integration is invisible regardless.
    /// </summary>
    public bool EnableOutlook { get; set; } = true;

    /// <summary>
    /// Copy every captured entry to Outlook without typing <c>!OL</c>. A line carrying
    /// <c>!NOL</c> still stays local. Off by default: sending work to another application
    /// should be something the user turns on knowingly.
    /// </summary>
    public bool PushToOutlookByDefault { get; set; }

    /// <summary>
    /// Draw the Outlook calendar under the user's own entries — meetings other people
    /// booked, rooms, everything Flowdeck did not put there. Read-only and never written
    /// back. Off by default: reading somebody's whole work diary is not a thing to assume.
    /// </summary>
    public bool ShowOutlookCalendar { get; set; }

    /// <summary>
    /// How often to re-read it, in minutes. Each read attaches to Outlook, so a short
    /// interval on a machine where Outlook is closed means waking it that often.
    /// </summary>
    public int OutlookCalendarRefreshMinutes { get; set; } = 10;

    /// <summary>
    /// A prompt when a meeting ends, asking for a line on it. On for every meeting by
    /// default: a note is worth most on the meeting nobody planned to take one on, and a
    /// balloon that goes away by itself is a small price for catching it.
    /// </summary>
    public MeetingNudgeScope MeetingNudge { get; set; } = MeetingNudgeScope.All;

    /// <summary>The interval as it will actually be used, whatever was typed into the box.</summary>
    [JsonIgnore]
    public int EffectiveOutlookRefreshMinutes => Math.Clamp(OutlookCalendarRefreshMinutes, 1, 1440);

    /// <summary>
    /// Whether to take entries that another application leaves in the inbox folder.
    ///
    /// On by default. What is watched is a folder inside Flowdeck's own data folder, so the
    /// only person who can put anything there is the one running the app; and somebody who
    /// has set up the sending side has already said what they want — making them find a
    /// second switch over here would just make the hand-over look broken.
    /// </summary>
    public bool EnableInboxWatch { get; set; } = true;

    /// <summary>
    /// Where that folder is. Null means <c>inbox</c> inside the data folder, which is what
    /// the sending side assumes; this is the way out for a machine where the data folder is
    /// a slow roaming profile, and both sides have to be pointed at the new place by hand.
    /// </summary>
    public string? InboxFolder { get; set; }

    /// <summary>How many tag colours there are before they come round again.</summary>
    public const int TagColorCount = 10;

    /// <summary>
    /// The colour each tag wears, as an index into the palette. Filled in as tags are first
    /// seen, in that order, and kept — a tag's colour has to mean the same thing tomorrow,
    /// which is the one thing the pair bars deliberately do not promise. Editable here by
    /// hand for a tag that has to be a particular colour.
    /// </summary>
    public Dictionary<string, int> TagColors { get; set; } = new();

    /// <summary>
    /// The colour for a tag, assigning the next one when the tag is new. The caller saves
    /// when <paramref name="assigned"/> comes back true; an assignment that is not written
    /// down would be a different colour next time, which defeats the purpose.
    /// </summary>
    public int TagColor(string tag, out bool assigned)
    {
        assigned = false;
        if (string.IsNullOrWhiteSpace(tag)) return -1;

        // The dictionary comes back from JSON with the default comparer, so the match is
        // done by hand: "#Motor" and "#motor" are one tag here as everywhere else.
        foreach (var (name, index) in TagColors)
        {
            if (string.Equals(name, tag, StringComparison.OrdinalIgnoreCase))
            {
                return ((index % TagColorCount) + TagColorCount) % TagColorCount;
            }
        }

        var next = TagColors.Count % TagColorCount;
        TagColors[tag] = next;
        assigned = true;
        return next;
    }

    /// <summary>Editable so the routing vocabulary can be tuned without a rebuild.</summary>
    public RoutingRules Routing { get; set; } = new();

    /// <summary>
    /// Names for <c>@handle</c> to resolve through, kept by hand in settings.json. There is
    /// no UI on purpose: it is a list you write once and forget, and a directory lookup
    /// would drag in exactly the permissions this app avoids needing.
    /// </summary>
    public ContactBook Contacts { get; set; } = new();

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

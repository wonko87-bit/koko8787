using System.Globalization;
using Flowdeck.Core.Models;

namespace Flowdeck.Core.Parsing;

/// <summary>
/// The structured reading of one line of input. This is what the quick-add window
/// previews live while the user types, and what <c>EntryService</c> turns into records.
/// </summary>
public sealed class ParsedEntry
{
    public const string UntitledPlaceholder = "(제목 없음)";

    public string RawInput { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public EntryTarget Target { get; set; } = EntryTarget.Both;

    /// <summary>True when the target came from an explicit <c>!TD</c>/<c>!CD</c> marker.</summary>
    public bool TargetWasExplicit { get; set; }

    /// <summary>
    /// Copy this entry to Outlook as well. Independent of <see cref="Target"/>, which
    /// decides whether that means an appointment, a task, or both.
    /// </summary>
    public bool PushExternal { get; set; }

    /// <summary>
    /// Open the Teams new-meeting form with this entry filled in. Nothing is sent: Teams
    /// mints the meeting only when the user presses its own save button.
    /// </summary>
    public bool OpenTeamsMeeting { get; set; }

    /// <summary>Email addresses pulled out of the line, in the order they were written.</summary>
    public List<string> Attendees { get; set; } = new();

    public DateTime? Start { get; set; }

    public DateTime? End { get; set; }

    /// <summary>True when a clock time was found, not just a calendar day.</summary>
    public bool HasTime { get; set; }

    public bool HasDate => Start.HasValue;

    public bool IsAllDay => HasDate && !HasTime;

    public Priority Priority { get; set; } = Priority.None;

    public List<string> Tags { get; set; } = new();

    public Recurrence Recurrence { get; set; } = Recurrence.None;

    public int? ReminderMinutesBefore { get; set; }

    /// <summary>True when the line carried nothing usable.</summary>
    public bool IsEmpty => string.IsNullOrWhiteSpace(RawInput);

    /// <summary>A one-line summary of what will be saved, for the live preview.</summary>
    public string DescribeSchedule()
    {
        if (!Start.HasValue) return "날짜 없음";

        var culture = CultureInfo.GetCultureInfo("ko-KR");
        var start = Start.Value;

        if (!HasTime)
        {
            var startLabel = start.ToString("M월 d일 (ddd)", culture);
            if (End.HasValue && End.Value.Date > start.Date)
            {
                return $"{startLabel} ~ {End.Value.ToString("M월 d일 (ddd)", culture)} · 종일";
            }

            return $"{startLabel} · 종일";
        }

        var dayLabel = start.ToString("M월 d일 (ddd)", culture);
        var startTime = start.ToString("tt h:mm", culture);

        if (!End.HasValue) return $"{dayLabel} {startTime}";

        var end = End.Value;
        return end.Date == start.Date
            ? $"{dayLabel} {startTime} ~ {end.ToString("tt h:mm", culture)}"
            : $"{dayLabel} {startTime} ~ {end.ToString("M월 d일 (ddd) tt h:mm", culture)}";
    }

    public string DescribeTarget() => Target switch
    {
        EntryTarget.Calendar => "일정",
        EntryTarget.Todo => "할일",
        _ => "할일 + 일정",
    };
}

namespace Flowdeck.Core.Models;

/// <summary>
/// A repeat rule. <see cref="Interval"/> is the multiplier on <see cref="Kind"/>,
/// so "격주" (every other week) is <c>Weekly</c> with an interval of 2.
/// </summary>
public sealed class Recurrence
{
    public RecurrenceKind Kind { get; set; } = RecurrenceKind.None;

    /// <summary>Every N units of <see cref="Kind"/>. Always at least 1.</summary>
    public int Interval { get; set; } = 1;

    /// <summary>
    /// For <see cref="RecurrenceKind.Weekly"/>: which weekdays it lands on.
    /// Empty means "the same weekday as the start date".
    /// </summary>
    public List<DayOfWeek> DaysOfWeek { get; set; } = new();

    /// <summary>Optional last date the rule produces occurrences for.</summary>
    public DateTime? Until { get; set; }

    public static Recurrence None => new() { Kind = RecurrenceKind.None };

    public bool IsRepeating => Kind != RecurrenceKind.None;

    public Recurrence Clone() => new()
    {
        Kind = Kind,
        Interval = Interval,
        DaysOfWeek = new List<DayOfWeek>(DaysOfWeek),
        Until = Until,
    };

    /// <summary>Human readable Korean label, e.g. "격주 화요일".</summary>
    public string Describe()
    {
        if (!IsRepeating) return string.Empty;

        var everyN = Interval > 1 ? $"{Interval}" : string.Empty;
        var baseLabel = Kind switch
        {
            RecurrenceKind.Daily => Interval > 1 ? $"{Interval}일마다" : "매일",
            RecurrenceKind.Weekdays => "평일마다",
            RecurrenceKind.Weekly => Interval == 2 ? "격주" : Interval > 1 ? $"{everyN}주마다" : "매주",
            RecurrenceKind.Monthly => Interval > 1 ? $"{everyN}개월마다" : "매월",
            RecurrenceKind.Yearly => Interval > 1 ? $"{everyN}년마다" : "매년",
            _ => string.Empty,
        };

        if (Kind == RecurrenceKind.Weekly && DaysOfWeek.Count > 0)
        {
            var days = string.Join(", ", DaysOfWeek.Select(KoreanDayName));
            return $"{baseLabel} {days}";
        }

        return baseLabel;
    }

    public static string KoreanDayName(DayOfWeek d) => d switch
    {
        DayOfWeek.Monday => "월요일",
        DayOfWeek.Tuesday => "화요일",
        DayOfWeek.Wednesday => "수요일",
        DayOfWeek.Thursday => "목요일",
        DayOfWeek.Friday => "금요일",
        DayOfWeek.Saturday => "토요일",
        _ => "일요일",
    };
}

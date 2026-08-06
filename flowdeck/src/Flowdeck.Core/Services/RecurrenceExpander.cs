using Flowdeck.Core.Models;

namespace Flowdeck.Core.Services;

/// <summary>
/// Expands a repeat rule into concrete dates. Nothing is stored per occurrence:
/// the calendar view asks for the window it is about to draw and gets the dates back.
/// </summary>
public static class RecurrenceExpander
{
    /// <summary>Ceiling on how many occurrences a single expansion may produce.</summary>
    private const int MaxOccurrences = 2000;

    /// <summary>
    /// Every occurrence of <paramref name="start"/> that falls inside
    /// [<paramref name="from"/>, <paramref name="to"/>], inclusive of both ends by date.
    /// </summary>
    public static IEnumerable<DateTime> Occurrences(
        DateTime start,
        Recurrence? recurrence,
        DateTime from,
        DateTime to)
    {
        var rule = recurrence ?? Recurrence.None;
        var windowStart = from.Date;
        var windowEnd = to.Date;

        if (windowEnd < windowStart) yield break;

        if (!rule.IsRepeating)
        {
            if (start.Date >= windowStart && start.Date <= windowEnd) yield return start;
            yield break;
        }

        var limit = rule.Until.HasValue && rule.Until.Value.Date < windowEnd
            ? rule.Until.Value.Date
            : windowEnd;

        var produced = 0;
        foreach (var occurrence in Walk(start, rule, limit))
        {
            if (++produced > MaxOccurrences) yield break;
            if (occurrence.Date < windowStart) continue;
            yield return occurrence;
        }
    }

    /// <summary>The first occurrence strictly after <paramref name="after"/>, if the rule has one.</summary>
    public static DateTime? Next(DateTime start, Recurrence? recurrence, DateTime after)
    {
        var rule = recurrence ?? Recurrence.None;
        if (!rule.IsRepeating) return null;

        var horizon = after.Date.AddYears(5);
        var limit = rule.Until.HasValue && rule.Until.Value.Date < horizon ? rule.Until.Value.Date : horizon;

        foreach (var occurrence in Walk(start, rule, limit))
        {
            if (occurrence > after) return occurrence;
        }

        return null;
    }

    private static IEnumerable<DateTime> Walk(DateTime start, Recurrence rule, DateTime limit)
    {
        var interval = Math.Max(1, rule.Interval);
        var timeOfDay = start.TimeOfDay;
        var guard = 0;

        switch (rule.Kind)
        {
            case RecurrenceKind.Daily:
                for (var cursor = start; cursor.Date <= limit; cursor = cursor.AddDays(interval))
                {
                    if (++guard > MaxOccurrences) yield break;
                    yield return cursor;
                }

                break;

            case RecurrenceKind.Weekdays:
                for (var cursor = start; cursor.Date <= limit; cursor = cursor.AddDays(1))
                {
                    if (++guard > MaxOccurrences) yield break;
                    if (cursor.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) continue;
                    yield return cursor;
                }

                break;

            case RecurrenceKind.Weekly:
            {
                var days = rule.DaysOfWeek.Count > 0
                    ? rule.DaysOfWeek.Distinct().OrderBy(WeekIndex).ToList()
                    : new List<DayOfWeek> { start.DayOfWeek };

                var weekStart = StartOfWeek(start.Date);
                while (weekStart <= limit)
                {
                    foreach (var day in days)
                    {
                        var date = weekStart.AddDays(WeekIndex(day)) + timeOfDay;
                        if (date < start) continue;
                        if (date.Date > limit) continue;
                        if (++guard > MaxOccurrences) yield break;
                        yield return date;
                    }

                    weekStart = weekStart.AddDays(7 * interval);
                }

                break;
            }

            case RecurrenceKind.Monthly:
            {
                var dayOfMonth = start.Day;
                var anchor = new DateTime(start.Year, start.Month, 1);
                while (anchor.Date <= limit)
                {
                    var daysInMonth = DateTime.DaysInMonth(anchor.Year, anchor.Month);
                    var date = new DateTime(anchor.Year, anchor.Month, Math.Min(dayOfMonth, daysInMonth)) + timeOfDay;
                    if (date >= start && date.Date <= limit)
                    {
                        if (++guard > MaxOccurrences) yield break;
                        yield return date;
                    }

                    anchor = anchor.AddMonths(interval);
                }

                break;
            }

            case RecurrenceKind.Yearly:
                for (var cursor = start; cursor.Date <= limit; cursor = cursor.AddYears(interval))
                {
                    if (++guard > MaxOccurrences) yield break;
                    yield return cursor;
                }

                break;
        }
    }

    private static DateTime StartOfWeek(DateTime date) => date.AddDays(-WeekIndex(date.DayOfWeek));

    private static int WeekIndex(DayOfWeek day) => ((int)day + 6) % 7;
}

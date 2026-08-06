using System.Globalization;
using System.Text;
using Flowdeck.Core.Models;

namespace Flowdeck.Core.Export;

/// <summary>
/// Writes events out as an iCalendar (.ics) file.
///
/// This is the bridge to everything else: importing the file into Google Calendar or
/// Outlook puts the same entries on a phone without the app needing an account or a server.
/// </summary>
public static class IcsExporter
{
    private const string ProductId = "-//Flowdeck//Desktop//KO";

    public static string Export(IEnumerable<CalendarEvent> events, DateTime stamp)
    {
        var builder = new StringBuilder();
        builder.Append("BEGIN:VCALENDAR\r\n");
        builder.Append("VERSION:2.0\r\n");
        builder.Append("PRODID:").Append(ProductId).Append("\r\n");
        builder.Append("CALSCALE:GREGORIAN\r\n");

        foreach (var item in events)
        {
            AppendEvent(builder, item, stamp);
        }

        builder.Append("END:VCALENDAR\r\n");
        return builder.ToString();
    }

    private static void AppendEvent(StringBuilder builder, CalendarEvent item, DateTime stamp)
    {
        builder.Append("BEGIN:VEVENT\r\n");
        builder.Append("UID:").Append(item.Id).Append("@flowdeck\r\n");
        builder.Append("DTSTAMP:").Append(FormatUtc(stamp)).Append("\r\n");

        if (item.IsAllDay)
        {
            builder.Append("DTSTART;VALUE=DATE:").Append(FormatDate(item.Start)).Append("\r\n");
            // DTEND is exclusive for all-day events, so a single day ends on the next date.
            builder.Append("DTEND;VALUE=DATE:").Append(FormatDate(item.End.Date.AddDays(1))).Append("\r\n");
        }
        else
        {
            builder.Append("DTSTART:").Append(FormatLocal(item.Start)).Append("\r\n");
            builder.Append("DTEND:").Append(FormatLocal(item.End)).Append("\r\n");
        }

        builder.Append("SUMMARY:").Append(Escape(item.Title)).Append("\r\n");

        if (!string.IsNullOrWhiteSpace(item.Notes))
        {
            builder.Append("DESCRIPTION:").Append(Escape(item.Notes)).Append("\r\n");
        }

        if (!string.IsNullOrWhiteSpace(item.Location))
        {
            builder.Append("LOCATION:").Append(Escape(item.Location)).Append("\r\n");
        }

        if (item.Tags.Count > 0)
        {
            builder.Append("CATEGORIES:").Append(Escape(string.Join(",", item.Tags))).Append("\r\n");
        }

        var rule = BuildRecurrenceRule(item.Recurrence);
        if (rule is not null) builder.Append("RRULE:").Append(rule).Append("\r\n");

        if (item.ReminderMinutesBefore is > 0)
        {
            builder.Append("BEGIN:VALARM\r\n");
            builder.Append("ACTION:DISPLAY\r\n");
            builder.Append("DESCRIPTION:").Append(Escape(item.Title)).Append("\r\n");
            builder.Append("TRIGGER:-PT").Append(item.ReminderMinutesBefore.Value).Append("M\r\n");
            builder.Append("END:VALARM\r\n");
        }

        builder.Append("END:VEVENT\r\n");
    }

    private static string? BuildRecurrenceRule(Recurrence recurrence)
    {
        if (!recurrence.IsRepeating) return null;

        var interval = Math.Max(1, recurrence.Interval);
        var parts = new List<string>();

        switch (recurrence.Kind)
        {
            case RecurrenceKind.Daily:
                parts.Add("FREQ=DAILY");
                break;

            case RecurrenceKind.Weekdays:
                parts.Add("FREQ=WEEKLY");
                parts.Add("BYDAY=MO,TU,WE,TH,FR");
                break;

            case RecurrenceKind.Weekly:
                parts.Add("FREQ=WEEKLY");
                if (recurrence.DaysOfWeek.Count > 0)
                {
                    parts.Add("BYDAY=" + string.Join(",", recurrence.DaysOfWeek.Select(IcsDay)));
                }

                break;

            case RecurrenceKind.Monthly:
                parts.Add("FREQ=MONTHLY");
                break;

            case RecurrenceKind.Yearly:
                parts.Add("FREQ=YEARLY");
                break;

            default:
                return null;
        }

        if (interval > 1) parts.Add("INTERVAL=" + interval.ToString(CultureInfo.InvariantCulture));
        if (recurrence.Until.HasValue) parts.Add("UNTIL=" + FormatUtc(recurrence.Until.Value.ToUniversalTime()));

        return string.Join(";", parts);
    }

    private static string IcsDay(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => "MO",
        DayOfWeek.Tuesday => "TU",
        DayOfWeek.Wednesday => "WE",
        DayOfWeek.Thursday => "TH",
        DayOfWeek.Friday => "FR",
        DayOfWeek.Saturday => "SA",
        _ => "SU",
    };

    private static string FormatDate(DateTime value) => value.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

    private static string FormatLocal(DateTime value) =>
        value.ToString("yyyyMMdd'T'HHmmss", CultureInfo.InvariantCulture);

    private static string FormatUtc(DateTime value) =>
        value.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);

    /// <summary>Escapes the characters iCalendar treats as structural.</summary>
    private static string Escape(string value) => value
        .Replace("\\", "\\\\")
        .Replace(";", "\\;")
        .Replace(",", "\\,")
        .Replace("\r\n", "\\n")
        .Replace("\n", "\\n");
}

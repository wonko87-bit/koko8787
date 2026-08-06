namespace Flowdeck.Core.Models;

/// <summary>A block of time on the calendar.</summary>
public sealed class CalendarEvent
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Title { get; set; } = string.Empty;

    public string Notes { get; set; } = string.Empty;

    public string Location { get; set; } = string.Empty;

    public DateTime Start { get; set; }

    public DateTime End { get; set; }

    /// <summary>True when the input carried a date but no clock time.</summary>
    public bool IsAllDay { get; set; }

    public List<string> Tags { get; set; } = new();

    public Recurrence Recurrence { get; set; } = Recurrence.None;

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    public DateTime UpdatedAt { get; set; } = DateTime.Now;

    /// <summary>Set when the same input also produced a todo item.</summary>
    public string? LinkedTodoId { get; set; }

    public string SourceText { get; set; } = string.Empty;

    public int? ReminderMinutesBefore { get; set; }

    /// <summary>True when this event covers any part of <paramref name="day"/>.</summary>
    public bool OccursOn(DateTime day)
    {
        var dayStart = day.Date;
        var dayEnd = dayStart.AddDays(1);
        return Start < dayEnd && End > dayStart
               || (IsAllDay && Start.Date <= dayStart && End.Date >= dayStart);
    }

    public CalendarEvent Clone() => (CalendarEvent)MemberwiseClone();
}

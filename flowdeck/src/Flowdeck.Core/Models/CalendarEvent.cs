namespace Flowdeck.Core.Models;

/// <summary>
/// Where a copied-in meeting came from: one occurrence of one appointment in the outside
/// calendar. The start is part of the identity because a repeating meeting shares one entry
/// id across every morning it falls on, and taking a copy of this Monday's must not make
/// every other Monday disappear from the overlay.
/// </summary>
public sealed record ExternalOrigin(string EntryId, DateTime Start);

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

    /// <summary>Set once the record has been copied to Outlook or another external store.</summary>
    public ExternalLink? ExternalLink { get; set; }

    /// <summary>
    /// Set when this is the user's own copy of a meeting read from the outside calendar —
    /// somebody else's booking, taken in so a note can be kept on it.
    ///
    /// The opposite of <see cref="ExternalLink"/>. A link is a copy Flowdeck made over there
    /// and keeps in step; an origin is a copy Flowdeck made here and never writes back. The
    /// meeting is not the user's to change, and the note is not the meeting's to carry.
    /// </summary>
    public ExternalOrigin? Origin { get; set; }

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

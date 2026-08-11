namespace Flowdeck.Core.Integration;

/// <summary>
/// One appointment as it stands in the outside calendar, already expanded — a repeating
/// meeting arrives as the individual mornings it actually falls on, not as a rule.
///
/// Read-only, and never stored. It exists to be drawn next to the user's own entries and
/// then thrown away on the next read, which is what keeps the whole feature free of the
/// question of who wins when both sides change.
/// </summary>
public sealed record ExternalOccurrence(
    string EntryId,
    string Title,
    DateTime Start,
    DateTime End,
    bool IsAllDay,
    string Location)
{
    /// <summary>Every date this occurrence touches, so a multi-day meeting shows on each.</summary>
    public IEnumerable<DateTime> Days()
    {
        var last = End <= Start ? Start.Date : End.AddTicks(-1).Date;
        for (var day = Start.Date; day <= last; day = day.AddDays(1)) yield return day;
    }
}

/// <summary>
/// Reads the calendar Flowdeck writes to, so the work diary can be shown alongside what was
/// typed here.
///
/// Separate from <see cref="IExternalStore"/> on purpose. Writing is about the entries the
/// user made; this is about everything else on their calendar — meetings other people
/// booked, rooms, the things Flowdeck knows nothing about. A store may do one and not the
/// other, and the app checks for this interface before offering the feature.
/// </summary>
public interface IExternalCalendarReader
{
    /// <summary>Shown in settings, e.g. "Outlook".</summary>
    string DisplayName { get; }

    /// <summary>
    /// Everything on the calendar that overlaps the window, with repeats already expanded
    /// into occurrences. Throws if the calendar could not be read at all.
    /// </summary>
    Task<IReadOnlyList<ExternalOccurrence>> ReadCalendarAsync(
        DateTime from,
        DateTime to,
        CancellationToken cancellationToken = default);
}

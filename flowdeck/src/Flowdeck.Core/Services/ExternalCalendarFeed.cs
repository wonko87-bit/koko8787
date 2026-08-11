using Flowdeck.Core.Integration;

namespace Flowdeck.Core.Services;

/// <summary>
/// The work calendar, held in memory and refreshed on a timer, to be drawn underneath the
/// user's own entries.
///
/// Strictly read-only and never persisted. Nothing here is ever written back, which is what
/// makes the whole thing safe: there is no version of this feature where a stale read
/// overwrites a real meeting. Losing the lot on restart costs one refresh.
/// </summary>
public sealed class ExternalCalendarFeed
{
    private readonly IExternalCalendarReader? _reader;
    private readonly Func<DateTime> _clock;

    private IReadOnlyList<ExternalOccurrence> _occurrences = Array.Empty<ExternalOccurrence>();
    private Dictionary<DateTime, List<ExternalOccurrence>> _byDay = new();

    private DateTime _from = DateTime.MinValue;
    private DateTime _to = DateTime.MinValue;

    public ExternalCalendarFeed(IExternalCalendarReader? reader, Func<DateTime>? clock = null)
    {
        _reader = reader;
        _clock = clock ?? (() => DateTime.Now);
    }

    /// <summary>Raised after a refresh changed what is on screen.</summary>
    public event EventHandler? Changed;

    /// <summary>False when there is nothing to read from — no Outlook, or a store that only writes.</summary>
    public bool IsAvailable => _reader is not null;

    public string DisplayName => _reader?.DisplayName ?? "Outlook";

    /// <summary>Off until the user asks for it. Reading someone's whole diary is not a default.</summary>
    public bool IsEnabled { get; set; }

    /// <summary>What went wrong on the last attempt, for the settings screen to show.</summary>
    public string? LastError { get; private set; }

    public DateTime? LastReadAt { get; private set; }

    public int Count => _occurrences.Count;

    /// <summary>
    /// The occurrences on one day, minus the ones Flowdeck put there itself.
    ///
    /// Without that subtraction every entry sent with <c>!OL</c> would come straight back
    /// and appear twice — once as the user's own row, once as a meeting read from Outlook.
    /// </summary>
    public IReadOnlyList<ExternalOccurrence> On(DateTime day, IReadOnlySet<string> exclude)
    {
        if (!IsEnabled || !_byDay.TryGetValue(day.Date, out var found)) return Array.Empty<ExternalOccurrence>();

        return found.Where(o => !exclude.Contains(o.EntryId)).ToList();
    }

    /// <summary>Days in the window carrying at least one meeting of their own, for the month dots.</summary>
    public HashSet<DateTime> DaysWith(IReadOnlySet<string> exclude)
    {
        var days = new HashSet<DateTime>();
        if (!IsEnabled) return days;

        foreach (var occurrence in _occurrences)
        {
            if (exclude.Contains(occurrence.EntryId)) continue;
            foreach (var day in occurrence.Days()) days.Add(day);
        }

        return days;
    }

    /// <summary>True when the window already covers these dates, so paging does not re-read.</summary>
    public bool Covers(DateTime from, DateTime to) => _from <= from.Date && _to >= to.Date;

    /// <summary>
    /// Reads the window again. Never throws: a mailbox that will not answer leaves the last
    /// good read on screen and puts the reason in <see cref="LastError"/>, because a calendar
    /// that empties itself every time Outlook is busy is worse than one that is a bit stale.
    ///
    /// A read that found the same diary as last time changes nothing and says nothing. On a
    /// short interval that is nearly every read, and each announcement costs the widget a
    /// full rebuild — the month grid, the day's rows, all of it — for a screen that would
    /// have looked identical. Most days a work calendar sits still for hours at a time.
    /// </summary>
    public async Task RefreshAsync(DateTime from, DateTime to, CancellationToken cancellationToken = default)
    {
        if (_reader is null || !IsEnabled) return;

        bool changed;

        try
        {
            var read = await _reader.ReadCalendarAsync(from.Date, to.Date.AddDays(1).AddTicks(-1), cancellationToken);

            // Occurrences are records read in start order, so comparing them in sequence is
            // both cheap and exact — a moved meeting or a changed room fails it.
            changed = !read.SequenceEqual(_occurrences);

            if (changed)
            {
                _occurrences = read;
                _byDay = Group(read);
            }

            _from = from.Date;
            _to = to.Date;
            LastError = null;
            LastReadAt = _clock();
        }
        catch (Exception e)
        {
            LastError = e.Message;
            return;
        }

        if (changed) Changed?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Drops everything, for when the user switches the overlay off.</summary>
    public void Clear()
    {
        if (_occurrences.Count == 0) return;

        _occurrences = Array.Empty<ExternalOccurrence>();
        _byDay = new Dictionary<DateTime, List<ExternalOccurrence>>();
        _from = DateTime.MinValue;
        _to = DateTime.MinValue;
        Changed?.Invoke(this, EventArgs.Empty);
    }

    private static Dictionary<DateTime, List<ExternalOccurrence>> Group(IReadOnlyList<ExternalOccurrence> read)
    {
        var byDay = new Dictionary<DateTime, List<ExternalOccurrence>>();

        foreach (var occurrence in read)
        {
            foreach (var day in occurrence.Days())
            {
                if (!byDay.TryGetValue(day, out var list)) byDay[day] = list = new List<ExternalOccurrence>();
                list.Add(occurrence);
            }
        }

        foreach (var list in byDay.Values)
        {
            list.Sort((a, b) =>
            {
                // All-day first, then by start, so a day reads top to bottom.
                if (a.IsAllDay != b.IsAllDay) return a.IsAllDay ? -1 : 1;
                return a.Start.CompareTo(b.Start);
            });
        }

        return byDay;
    }
}

using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Storage;

namespace Flowdeck.Core.Services;

/// <summary>A tag and how many entries carry it.</summary>
public sealed record TagCount(string Tag, int Count);

/// <summary>One occurrence of an event, with the repeat rule already applied.</summary>
public sealed class EventOccurrence
{
    public required CalendarEvent Source { get; init; }

    public DateTime Start { get; init; }

    public DateTime End { get; init; }

    public string Title => Source.Title;

    public bool IsAllDay => Source.IsAllDay;

    public bool IsRecurring => Source.Recurrence.IsRepeating;
}

/// <summary>
/// The in-memory workspace plus the operations the UI performs on it.
/// Every mutation writes through to the store and raises <see cref="Changed"/>.
/// </summary>
public sealed class WorkspaceRepository
{
    private readonly IWorkspaceStore _store;
    private Workspace _workspace = new();

    public WorkspaceRepository(IWorkspaceStore store) => _store = store;

    /// <summary>Raised after any mutation, on the thread that performed it.</summary>
    public event EventHandler? Changed;

    public IReadOnlyList<TodoItem> Todos => _workspace.Todos;

    public IReadOnlyList<CalendarEvent> Events => _workspace.Events;

    /// <summary>
    /// Note the deliberate absence of ConfigureAwait(false) here and in every mutation:
    /// <see cref="Changed"/> has to reach the caller's context, because on the desktop the
    /// handler updates collections that are bound to the UI and may only be touched there.
    /// </summary>
    public async Task LoadAsync(CancellationToken cancellationToken = default)
    {
        _workspace = await _store.LoadAsync(cancellationToken);
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public Task SaveAsync(CancellationToken cancellationToken = default) =>
        _store.SaveAsync(_workspace, cancellationToken);

    /// <summary>Files a parsed line and persists the result.</summary>
    public async Task<CaptureResult> CaptureAsync(ParsedEntry entry, DateTime now)
    {
        var result = EntryComposer.Compose(entry, now);
        if (result.IsEmpty) return result;

        if (result.Todo is not null) _workspace.Todos.Add(result.Todo);
        if (result.Event is not null) _workspace.Events.Add(result.Event);

        await SaveAsync();
        Changed?.Invoke(this, EventArgs.Empty);
        return result;
    }

    /// <summary>
    /// Ticks a todo off. A repeating todo is not closed: its due date advances to the
    /// next occurrence, which is what "매일 아침 약 먹기" is supposed to do.
    /// </summary>
    public async Task ToggleTodoAsync(string id, DateTime now)
    {
        var todo = _workspace.Todos.FirstOrDefault(t => t.Id == id);
        if (todo is null) return;

        if (!todo.IsDone && todo.Recurrence.IsRepeating && todo.DueAt.HasValue)
        {
            var next = RecurrenceExpander.Next(todo.DueAt.Value, todo.Recurrence, todo.DueAt.Value);
            if (next.HasValue)
            {
                todo.DueAt = next.Value;
                todo.UpdatedAt = now;
                await SaveAsync().ConfigureAwait(false);
                Changed?.Invoke(this, EventArgs.Empty);
                return;
            }
        }

        todo.IsDone = !todo.IsDone;
        todo.CompletedAt = todo.IsDone ? now : null;
        todo.UpdatedAt = now;

        await SaveAsync();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public async Task DeleteTodoAsync(string id)
    {
        var todo = _workspace.Todos.FirstOrDefault(t => t.Id == id);
        if (todo is null) return;

        _workspace.Todos.Remove(todo);
        foreach (var linked in _workspace.Events.Where(e => e.LinkedTodoId == id))
        {
            linked.LinkedTodoId = null;
        }

        await SaveAsync();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public async Task DeleteEventAsync(string id)
    {
        var calendarEvent = _workspace.Events.FirstOrDefault(e => e.Id == id);
        if (calendarEvent is null) return;

        _workspace.Events.Remove(calendarEvent);
        foreach (var linked in _workspace.Todos.Where(t => t.LinkedEventId == id))
        {
            linked.LinkedEventId = null;
        }

        await SaveAsync();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Removes items completed before <paramref name="cutoff"/>, so the list stays short.</summary>
    public async Task<int> PurgeCompletedAsync(DateTime cutoff)
    {
        var removed = _workspace.Todos.RemoveAll(t => t.IsDone && t.CompletedAt.HasValue && t.CompletedAt < cutoff);
        if (removed == 0) return 0;

        await SaveAsync();
        Changed?.Invoke(this, EventArgs.Empty);
        return removed;
    }

    /// <summary>Every occurrence touching <paramref name="day"/>, ordered by start time.</summary>
    public IReadOnlyList<EventOccurrence> OccurrencesOn(DateTime day) =>
        OccurrencesBetween(day.Date, day.Date);

    /// <summary>Every occurrence between two dates, inclusive.</summary>
    public IReadOnlyList<EventOccurrence> OccurrencesBetween(DateTime from, DateTime to)
    {
        var results = new List<EventOccurrence>();

        foreach (var source in _workspace.Events)
        {
            var span = source.End - source.Start;
            if (span < TimeSpan.Zero) span = TimeSpan.Zero;

            foreach (var start in RecurrenceExpander.Occurrences(source.Start, source.Recurrence, from, to))
            {
                results.Add(new EventOccurrence { Source = source, Start = start, End = start + span });
            }
        }

        return results
            .OrderBy(o => o.IsAllDay ? 0 : 1)
            .ThenBy(o => o.Start)
            .ToList();
    }

    /// <summary>
    /// Events that happen exactly once, in date order and regardless of how far off they
    /// are — nothing a user typed should quietly fall off the list.
    ///
    /// Repeating events are excluded on purpose: they have no last occurrence, so listing
    /// them by date would bury the one-offs. <see cref="RepeatingEvents"/> returns those as
    /// rules instead.
    /// </summary>
    public IReadOnlyList<EventOccurrence> OneOffOccurrences()
    {
        var results = new List<EventOccurrence>();

        foreach (var source in _workspace.Events.Where(e => !e.Recurrence.IsRepeating))
        {
            var span = source.End - source.Start;
            if (span < TimeSpan.Zero) span = TimeSpan.Zero;

            results.Add(new EventOccurrence
            {
                Source = source,
                Start = source.Start,
                End = source.Start + span,
            });
        }

        return results.OrderBy(o => o.Start).ToList();
    }

    /// <summary>The repeat rules themselves, one entry per event, earliest start first.</summary>
    public IReadOnlyList<CalendarEvent> RepeatingEvents() =>
        _workspace.Events
            .Where(e => e.Recurrence.IsRepeating)
            .OrderBy(e => e.Start)
            .ToList();

    /// <summary>
    /// Every todo, oldest due date first, with undated items last. Completed items are
    /// left out unless asked for.
    /// </summary>
    public IReadOnlyList<TodoItem> AllTodos(bool includeCompleted) =>
        _workspace.Todos
            .Where(t => includeCompleted || !t.IsDone)
            .OrderBy(t => t.DueAt.HasValue ? 0 : 1)
            .ThenBy(t => t.DueAt ?? DateTime.MaxValue)
            .ThenByDescending(t => t.Priority)
            .ToList();

    /// <summary>
    /// Tags in use on the calendar, most used first. A repeating event counts once — it is
    /// one entry, however many times it comes round.
    /// </summary>
    public IReadOnlyList<TagCount> EventTags() => CountTags(_workspace.Events.Select(e => e.Tags));

    /// <summary>Tags in use on the todo list, most used first.</summary>
    public IReadOnlyList<TagCount> TodoTags(bool includeCompleted) =>
        CountTags(_workspace.Todos.Where(t => includeCompleted || !t.IsDone).Select(t => t.Tags));

    /// <summary>
    /// Groups case-insensitively so "#Work" and "#work" are one tag, keeping whichever
    /// spelling was seen first for display.
    /// </summary>
    private static IReadOnlyList<TagCount> CountTags(IEnumerable<IEnumerable<string>> tagLists)
    {
        var counts = new Dictionary<string, TagCount>(StringComparer.OrdinalIgnoreCase);

        foreach (var tag in tagLists.SelectMany(tags => tags))
        {
            counts[tag] = counts.TryGetValue(tag, out var seen)
                ? seen with { Count = seen.Count + 1 }
                : new TagCount(tag, 1);
        }

        return counts.Values
            .OrderByDescending(t => t.Count)
            .ThenBy(t => t.Tag, StringComparer.CurrentCulture)
            .ToList();
    }

    /// <summary>Which days in a range have at least one event, for dotting the month grid.</summary>
    public HashSet<DateTime> DaysWithEvents(DateTime from, DateTime to) =>
        OccurrencesBetween(from, to).Select(o => o.Start.Date).ToHashSet();

    public IReadOnlyList<TodoItem> OpenTodos() =>
        _workspace.Todos.Where(t => !t.IsDone).ToList();

    /// <summary>
    /// The working list for the widget: overdue first, then today, then everything
    /// else by due date, with undated items last.
    /// </summary>
    public IReadOnlyList<TodoItem> TodosForToday(DateTime now, bool includeUndated = true)
    {
        var endOfDay = now.Date.AddDays(1);

        return _workspace.Todos
            .Where(t => !t.IsDone)
            .Where(t => t.DueAt is null ? includeUndated : t.DueAt < endOfDay)
            .OrderBy(t => t.DueAt.HasValue ? 0 : 1)
            .ThenBy(t => t.DueAt ?? DateTime.MaxValue)
            .ThenByDescending(t => t.Priority)
            .ToList();
    }

    public IReadOnlyList<TodoItem> TodosOn(DateTime day) =>
        _workspace.Todos
            .Where(t => t.DueAt.HasValue && t.DueAt.Value.Date == day.Date)
            .OrderBy(t => t.IsDone)
            .ThenBy(t => t.DueAt)
            .ToList();

    /// <summary>Replaces the whole workspace, e.g. after an import.</summary>
    public async Task ReplaceAsync(Workspace workspace)
    {
        _workspace = workspace;
        await SaveAsync();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public Workspace Snapshot() => _workspace;
}

using System.Collections.ObjectModel;
using System.Windows.Threading;
using Flowdeck.Core.Services;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>Which of the two list windows this is.</summary>
public enum AgendaMode
{
    Events,
    Todos,
}

/// <summary>One day's worth of entries, under a single date heading.</summary>
public sealed class AgendaGroup
{
    public AgendaGroup(string header, bool isToday, bool isPast, bool isUndated)
    {
        Header = header;
        IsToday = isToday;
        IsPast = isPast;
        IsUndated = isUndated;
    }

    public string Header { get; }

    public bool IsToday { get; }

    /// <summary>Drawn back a step, so the eye lands on today and what is still ahead.</summary>
    public bool IsPast { get; }

    public bool IsUndated { get; }

    /// <summary>Holds <see cref="EventRow"/> or <see cref="TodoRow"/>; the view picks the template by type.</summary>
    public ObservableCollection<object> Items { get; } = new();

    public string CountLabel => Items.Count.ToString(Ko.Culture);
}

/// <summary>
/// Backs the standalone list windows: everything accumulated so far, in date order,
/// broken into one group per day.
/// </summary>
public sealed class AgendaViewModel : ObservableObject
{
    private readonly WorkspaceRepository _repository;
    private readonly Func<DateTime> _clock;
    private readonly Dispatcher? _dispatcher;

    private bool _showCompleted;
    private bool _isEmpty = true;

    public AgendaViewModel(WorkspaceRepository repository, AgendaMode mode, Func<DateTime>? clock = null)
    {
        _repository = repository;
        Mode = mode;
        _clock = clock ?? (() => DateTime.Now);

        _dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
        _repository.Changed += (_, _) => OnRepositoryChanged();

        Reload();
    }

    public AgendaMode Mode { get; }

    public ObservableCollection<AgendaGroup> Groups { get; } = new();

    /// <summary>Repeat rules, one line each, shown below the dated groups.</summary>
    public ObservableCollection<RecurringRow> Recurring { get; } = new();

    public bool HasRecurring => Recurring.Count > 0;

    public string RecurringSectionLabel => Mode == AgendaMode.Events ? "반복 일정" : "반복 할일";

    public string Title => Mode == AgendaMode.Events ? "일정" : "할일";

    public bool IsTodoMode => Mode == AgendaMode.Todos;

    /// <summary>Only meaningful for the todo window; the calendar has nothing to complete.</summary>
    public bool ShowCompleted
    {
        get => _showCompleted;
        set
        {
            if (!Set(ref _showCompleted, value)) return;
            Reload();
        }
    }

    public bool IsEmpty
    {
        get => _isEmpty;
        private set => Set(ref _isEmpty, value);
    }

    public string EmptyMessage => Mode == AgendaMode.Events ? "등록된 일정이 없습니다" : "등록된 할일이 없습니다";

    public void Reload()
    {
        var now = _clock();

        Groups.Clear();
        Recurring.Clear();

        if (Mode == AgendaMode.Events)
        {
            LoadEvents(now);
        }
        else
        {
            LoadTodos(now);
        }

        IsEmpty = Groups.Count == 0 && Recurring.Count == 0;
        Raise(nameof(HasRecurring));
        Raise(nameof(TotalLabel));
    }

    public string TotalLabel => $"{Groups.Sum(g => g.Items.Count) + Recurring.Count}건";

    private void LoadEvents(DateTime now)
    {
        foreach (var byDay in _repository.OneOffOccurrences().GroupBy(o => o.Start.Date).OrderBy(g => g.Key))
        {
            var group = CreateGroup(byDay.Key, now);
            foreach (var occurrence in byDay.OrderBy(o => o.IsAllDay ? 0 : 1).ThenBy(o => o.Start))
            {
                group.Items.Add(new EventRow(occurrence, DeleteEventAsync));
            }

            Groups.Add(group);
        }

        foreach (var source in _repository.RepeatingEvents())
        {
            Recurring.Add(new RecurringRow(source, now, DeleteRecurringEventAsync));
        }
    }

    private void LoadTodos(DateTime now)
    {
        var todos = _repository.AllTodos(_showCompleted);

        // Undated items have no day to sit under, so they collect in a group of their own
        // at the end rather than being dropped or pinned to today.
        foreach (var byDay in todos
                     .Where(t => !t.Recurrence.IsRepeating && t.DueAt.HasValue)
                     .GroupBy(t => t.DueAt!.Value.Date)
                     .OrderBy(g => g.Key))
        {
            var group = CreateGroup(byDay.Key, now);
            foreach (var todo in byDay.OrderBy(t => t.IsDone).ThenBy(t => t.DueAt))
            {
                group.Items.Add(new TodoRow(todo, now, ToggleTodoAsync, DeleteTodoAsync));
            }

            Groups.Add(group);
        }

        var undated = todos.Where(t => !t.Recurrence.IsRepeating && !t.DueAt.HasValue).ToList();
        if (undated.Count > 0)
        {
            var tail = new AgendaGroup("날짜 없음", isToday: false, isPast: false, isUndated: true);
            foreach (var todo in undated)
            {
                tail.Items.Add(new TodoRow(todo, now, ToggleTodoAsync, DeleteTodoAsync));
            }

            Groups.Add(tail);
        }

        foreach (var todo in todos.Where(t => t.Recurrence.IsRepeating))
        {
            Recurring.Add(new RecurringRow(todo, ToggleRecurringTodoAsync, DeleteRecurringTodoAsync));
        }
    }

    private static AgendaGroup CreateGroup(DateTime day, DateTime now)
    {
        var today = now.Date;
        var dayLabel = day.Year == today.Year
            ? day.ToString("M월 d일 (ddd)", Ko.Culture)
            : day.ToString("yyyy년 M월 d일 (ddd)", Ko.Culture);

        var relative = (day - today).Days switch
        {
            0 => "오늘",
            1 => "내일",
            2 => "모레",
            -1 => "어제",
            _ => null,
        };

        var header = relative is null ? dayLabel : $"{relative} · {dayLabel}";
        return new AgendaGroup(header, isToday: day == today, isPast: day < today, isUndated: false);
    }

    private void OnRepositoryChanged()
    {
        if (_dispatcher is null || _dispatcher.CheckAccess())
        {
            Reload();
            return;
        }

        _dispatcher.Invoke(Reload);
    }

    private Task ToggleTodoAsync(TodoRow row) => _repository.ToggleTodoAsync(row.Id, _clock());

    private Task DeleteTodoAsync(TodoRow row) => _repository.DeleteTodoAsync(row.Id);

    private Task DeleteEventAsync(EventRow row) => _repository.DeleteEventAsync(row.Id);

    private Task ToggleRecurringTodoAsync(RecurringRow row) => _repository.ToggleTodoAsync(row.Id, _clock());

    private Task DeleteRecurringTodoAsync(RecurringRow row) => _repository.DeleteTodoAsync(row.Id);

    private Task DeleteRecurringEventAsync(RecurringRow row) => _repository.DeleteEventAsync(row.Id);
}

using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Threading;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>One tag in the filter row, with how many entries carry it.</summary>
public sealed class TagChip : ObservableObject
{
    private bool _isSelected;

    public TagChip(string label, int count, bool isSelected, Action<TagChip> onToggle)
    {
        Label = "#" + label;
        Tag = label;
        CountLabel = count.ToString(Ko.Culture);
        _isSelected = isSelected;
        ToggleCommand = new RelayCommand(() => onToggle(this));
    }

    /// <summary>The raw tag, without the leading hash.</summary>
    public string Tag { get; }

    public string Label { get; }

    public string CountLabel { get; }

    public bool IsSelected
    {
        get => _isSelected;
        set => Set(ref _isSelected, value);
    }

    public ICommand ToggleCommand { get; }
}

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

    private readonly HashSet<string> _selectedTags = new(StringComparer.OrdinalIgnoreCase);

    private bool _showCompleted;
    private bool _isEmpty = true;
    private bool _matchAllTags = true;

    public AgendaViewModel(WorkspaceRepository repository, AgendaMode mode, Func<DateTime>? clock = null)
    {
        _repository = repository;
        Mode = mode;
        _clock = clock ?? (() => DateTime.Now);

        ClearTagsCommand = new RelayCommand(ClearTags);
        ToggleMatchModeCommand = new RelayCommand(() => MatchAllTags = !MatchAllTags);

        _dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
        _repository.Changed += (_, _) => OnRepositoryChanged();

        Reload();
    }

    public AgendaMode Mode { get; }

    public ObservableCollection<AgendaGroup> Groups { get; } = new();

    /// <summary>Repeat rules, one line each, shown below the dated groups.</summary>
    public ObservableCollection<RecurringRow> Recurring { get; } = new();

    /// <summary>Every tag in use in this window, most used first.</summary>
    public ObservableCollection<TagChip> TagChips { get; } = new();

    public bool HasTagChips => TagChips.Count > 0;

    public bool HasTagFilter => _selectedTags.Count > 0;

    public ICommand ClearTagsCommand { get; }

    public ICommand ToggleMatchModeCommand { get; }

    /// <summary>
    /// With more than one tag picked: true narrows to entries carrying all of them,
    /// false widens to entries carrying any. Narrowing is the default, since that is
    /// what picking a second filter usually means.
    /// </summary>
    public bool MatchAllTags
    {
        get => _matchAllTags;
        set
        {
            if (!Set(ref _matchAllTags, value)) return;
            Raise(nameof(MatchModeLabel));
            if (_selectedTags.Count > 1) Reload();
        }
    }

    public string MatchModeLabel => _matchAllTags ? "모두 포함" : "하나라도";

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

    public string EmptyMessage
    {
        get
        {
            if (HasTagFilter)
            {
                return _matchAllTags && _selectedTags.Count > 1
                    ? "선택한 태그를 모두 가진 항목이 없습니다"
                    : "선택한 태그가 달린 항목이 없습니다";
            }

            return Mode == AgendaMode.Events ? "등록된 일정이 없습니다" : "등록된 할일이 없습니다";
        }
    }

    /// <summary>Drops the tag filter. Called when the window is summoned, so a filter set
    /// earlier never quietly hides entries the next time the list is opened.</summary>
    public void ClearTags()
    {
        if (_selectedTags.Count == 0) return;

        _selectedTags.Clear();
        Reload();
    }

    public void Reload()
    {
        var now = _clock();

        BuildTagChips();

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
        Raise(nameof(HasTagChips));
        Raise(nameof(HasTagFilter));
        Raise(nameof(EmptyMessage));
        Raise(nameof(TotalLabel));
    }

    private void BuildTagChips()
    {
        var counts = Mode == AgendaMode.Events
            ? _repository.EventTags()
            : _repository.TodoTags(_showCompleted);

        // A tag can disappear when its last entry is deleted; drop it from the selection
        // too, or the list would filter on something with no chip left to unselect it.
        _selectedTags.IntersectWith(counts.Select(c => c.Tag));

        TagChips.Clear();
        foreach (var count in counts)
        {
            TagChips.Add(new TagChip(count.Tag, count.Count, _selectedTags.Contains(count.Tag), OnTagToggled));
        }
    }

    private void OnTagToggled(TagChip chip)
    {
        if (!_selectedTags.Add(chip.Tag)) _selectedTags.Remove(chip.Tag);

        Reload();
    }

    /// <summary>Whether an entry's tags satisfy the current filter.</summary>
    private bool MatchesTagFilter(IReadOnlyList<string> tags)
    {
        if (_selectedTags.Count == 0) return true;

        return _matchAllTags
            ? _selectedTags.All(selected => tags.Contains(selected, StringComparer.OrdinalIgnoreCase))
            : tags.Any(tag => _selectedTags.Contains(tag));
    }

    public string TotalLabel => $"{Groups.Sum(g => g.Items.Count) + Recurring.Count}건";

    /// <summary>
    /// The entries behind what is on screen right now, repeating ones included.
    ///
    /// This is what makes the tag chips double as the picker for an export: narrowing the
    /// list narrows the file, with no second selection UI to learn and nothing that can
    /// disagree with what the user is looking at.
    /// </summary>
    public (IReadOnlyList<TodoItem> Todos, IReadOnlyList<CalendarEvent> Events) VisibleEntries()
    {
        if (Mode == AgendaMode.Events)
        {
            var events = _repository.OneOffOccurrences().Select(o => o.Source)
                .Concat(_repository.RepeatingEvents())
                .Where(e => MatchesTagFilter(e.Tags))
                .ToList();

            return (Array.Empty<TodoItem>(), events);
        }

        var todos = _repository.AllTodos(_showCompleted)
            .Where(t => MatchesTagFilter(t.Tags))
            .ToList();

        return (todos, Array.Empty<CalendarEvent>());
    }

    private void LoadEvents(DateTime now)
    {
        var oneOffs = _repository.OneOffOccurrences().Where(o => MatchesTagFilter(o.Source.Tags));

        foreach (var byDay in oneOffs.GroupBy(o => o.Start.Date).OrderBy(g => g.Key))
        {
            var group = CreateGroup(byDay.Key, now);
            foreach (var occurrence in byDay.OrderBy(o => o.IsAllDay ? 0 : 1).ThenBy(o => o.Start))
            {
                group.Items.Add(new EventRow(occurrence, DeleteEventAsync));
            }

            Groups.Add(group);
        }

        foreach (var source in _repository.RepeatingEvents().Where(e => MatchesTagFilter(e.Tags)))
        {
            Recurring.Add(new RecurringRow(source, now, DeleteRecurringEventAsync));
        }
    }

    private void LoadTodos(DateTime now)
    {
        var todos = _repository.AllTodos(_showCompleted)
            .Where(t => MatchesTagFilter(t.Tags))
            .ToList();

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

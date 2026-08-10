using System.Collections.ObjectModel;
using System.Windows.Input;
using Flowdeck.Core.Export;
using Flowdeck.Core.Models;
using Flowdeck.Mobile.Mvvm;
using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile.ViewModels;

/// <summary>
/// The list screen. One screen for both kinds, switched by a segmented control rather than
/// split across two, because a phone has no room for two windows side by side.
///
/// The tag chips double as the picker for an export, exactly as on the desktop: whatever
/// they have narrowed the list to is what goes into the file.
/// </summary>
public sealed class ListViewModel : ObservableObject
{
    private readonly Func<DateTime> _clock;
    private readonly HashSet<string> _selectedTags = new(StringComparer.OrdinalIgnoreCase);

    private bool _showEvents;
    private bool _showCompleted;
    private bool _isEmpty = true;
    private string _status = string.Empty;

    public ListViewModel(Func<DateTime>? clock = null)
    {
        _clock = clock ?? (() => DateTime.Now);

        ClearTagsCommand = new RelayCommand(ClearTags);
        ExportCommand = new AsyncRelayCommand(ExportAsync);
        ShowTodosCommand = new RelayCommand(() => ShowEvents = false);
        ShowEventsCommand = new RelayCommand(() => ShowEvents = true);

        Workspace.Repository.Changed += (_, _) => MainThread.BeginInvokeOnMainThread(Reload);
    }

    public ObservableCollection<DayGroup> Groups { get; } = new();

    public ObservableCollection<RecurringRow> Recurring { get; } = new();

    public ObservableCollection<TagChip> TagChips { get; } = new();

    public ICommand ClearTagsCommand { get; }

    public ICommand ExportCommand { get; }

    public ICommand ShowTodosCommand { get; }

    public ICommand ShowEventsCommand { get; }

    /// <summary>False shows todos, true shows the calendar.</summary>
    public bool ShowEvents
    {
        get => _showEvents;
        set
        {
            if (!Set(ref _showEvents, value)) return;

            // A tag that exists on one side may not exist on the other, so a filter carried
            // across would silently empty the list.
            _selectedTags.Clear();
            Raise(nameof(IsTodoMode));
            Raise(nameof(RecurringSectionLabel));
            Reload();
        }
    }

    public bool IsTodoMode => !_showEvents;

    public bool ShowCompleted
    {
        get => _showCompleted;
        set
        {
            if (Set(ref _showCompleted, value)) Reload();
        }
    }

    public bool HasTagChips => TagChips.Count > 0;

    public bool HasTagFilter => _selectedTags.Count > 0;

    public bool HasRecurring => Recurring.Count > 0;

    public string RecurringSectionLabel => _showEvents ? "반복 일정" : "반복 할일";

    public bool IsEmpty
    {
        get => _isEmpty;
        private set => Set(ref _isEmpty, value);
    }

    public string EmptyMessage => HasTagFilter
        ? "고른 태그에 해당하는 항목이 없습니다"
        : _showEvents
            ? "등록된 일정이 없습니다"
            : "등록된 할일이 없습니다";

    public string TotalLabel => $"{Groups.Sum(g => g.Count) + Recurring.Count}건";

    public string Status
    {
        get => _status;
        set
        {
            if (Set(ref _status, value)) Raise(nameof(HasStatus));
        }
    }

    public bool HasStatus => _status.Length > 0;

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

        if (_showEvents) LoadEvents(now); else LoadTodos(now);

        IsEmpty = Groups.Count == 0 && Recurring.Count == 0;
        Raise(nameof(HasRecurring));
        Raise(nameof(HasTagChips));
        Raise(nameof(HasTagFilter));
        Raise(nameof(EmptyMessage));
        Raise(nameof(TotalLabel));
    }

    /// <summary>
    /// The entries behind what is on screen, repeating ones included — the export payload.
    /// </summary>
    public (IReadOnlyList<TodoItem> Todos, IReadOnlyList<CalendarEvent> Events) VisibleEntries()
    {
        if (_showEvents)
        {
            var events = Workspace.Repository.OneOffOccurrences().Select(o => o.Source)
                .Concat(Workspace.Repository.RepeatingEvents())
                .Where(e => Matches(e.Tags))
                .ToList();

            return (Array.Empty<TodoItem>(), events);
        }

        var todos = Workspace.Repository.AllTodos(_showCompleted).Where(t => Matches(t.Tags)).ToList();
        return (todos, Array.Empty<CalendarEvent>());
    }

    /// <summary>
    /// Writes the visible entries out and hands the file to Android's share sheet, which is
    /// how anything leaves a phone. Where it goes from there — a chat to yourself, mail,
    /// a drive — is the user's business and the system remembers their habits better than
    /// this app could.
    /// </summary>
    public async Task ExportAsync()
    {
        var (todos, events) = VisibleEntries();
        var count = todos.Count + events.Count;
        if (count == 0)
        {
            Status = "내보낼 항목이 없습니다";
            return;
        }

        try
        {
            var name = $"flowdeck-{(_showEvents ? "일정" : "할일")}-{_clock():yyyyMMdd-HHmm}{TransferFile.Extension}";
            var path = Path.Combine(FileSystem.CacheDirectory, name);
            File.WriteAllText(path, TransferFile.Write(todos, events, DateTimeOffset.Now));

            await Share.Default.RequestAsync(new ShareFileRequest
            {
                Title = $"Flowdeck {count}건",
                File = new ShareFile(path),
            });

            Status = $"{count}건을 내보냈습니다";
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            Status = "내보내지 못했습니다: " + e.Message;
        }
    }

    private bool Matches(IReadOnlyList<string> tags)
    {
        if (_selectedTags.Count == 0) return true;

        // Any-of rather than all-of. On a phone the chips are thumbed at, and widening the
        // list is a cheaper mistake to recover from than emptying it.
        return tags.Any(tag => _selectedTags.Contains(tag));
    }

    private void BuildTagChips()
    {
        var counts = _showEvents
            ? Workspace.Repository.EventTags()
            : Workspace.Repository.TodoTags(_showCompleted);

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

    private void LoadTodos(DateTime now)
    {
        var todos = Workspace.Repository.AllTodos(_showCompleted).Where(t => Matches(t.Tags)).ToList();

        foreach (var byDay in todos
                     .Where(t => !t.Recurrence.IsRepeating && t.DueAt.HasValue)
                     .GroupBy(t => t.DueAt!.Value.Date)
                     .OrderBy(g => g.Key))
        {
            var group = CreateGroup(byDay.Key, now);
            foreach (var todo in byDay.OrderBy(t => t.IsDone).ThenBy(t => t.DueAt))
            {
                group.Add(new TodoRow(todo, now, ToggleTodoAsync, DeleteTodoAsync));
            }

            Groups.Add(group);
        }

        var undated = todos.Where(t => !t.Recurrence.IsRepeating && !t.DueAt.HasValue).ToList();
        if (undated.Count > 0)
        {
            var tail = new DayGroup("날짜 없음", isToday: false, isPast: false);
            foreach (var todo in undated) tail.Add(new TodoRow(todo, now, ToggleTodoAsync, DeleteTodoAsync));
            Groups.Add(tail);
        }

        foreach (var todo in todos.Where(t => t.Recurrence.IsRepeating))
        {
            Recurring.Add(new RecurringRow(todo, DeleteRecurringTodoAsync));
        }
    }

    private void LoadEvents(DateTime now)
    {
        var oneOffs = Workspace.Repository.OneOffOccurrences().Where(o => Matches(o.Source.Tags));

        foreach (var byDay in oneOffs.GroupBy(o => o.Start.Date).OrderBy(g => g.Key))
        {
            var group = CreateGroup(byDay.Key, now);
            foreach (var occurrence in byDay.OrderBy(o => o.IsAllDay ? 0 : 1).ThenBy(o => o.Start))
            {
                group.Add(new EventRow(occurrence, DeleteEventAsync));
            }

            Groups.Add(group);
        }

        foreach (var source in Workspace.Repository.RepeatingEvents().Where(e => Matches(e.Tags)))
        {
            Recurring.Add(new RecurringRow(source, DeleteRecurringEventAsync));
        }
    }

    private static DayGroup CreateGroup(DateTime day, DateTime now)
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

        return new DayGroup(
            relative is null ? dayLabel : $"{relative} · {dayLabel}",
            isToday: day == today,
            isPast: day < today);
    }

    private Task ToggleTodoAsync(TodoRow row) => Workspace.Repository.ToggleTodoAsync(row.Id, _clock());

    private Task DeleteTodoAsync(TodoRow row) => Workspace.Repository.DeleteTodoAsync(row.Id);

    private Task DeleteEventAsync(EventRow row) => Workspace.Repository.DeleteEventAsync(row.Id);

    private Task DeleteRecurringTodoAsync(RecurringRow row) => Workspace.Repository.DeleteTodoAsync(row.Id);

    private Task DeleteRecurringEventAsync(RecurringRow row) => Workspace.Repository.DeleteEventAsync(row.Id);
}

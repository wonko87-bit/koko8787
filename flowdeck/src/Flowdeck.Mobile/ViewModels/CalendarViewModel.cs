using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Input;
using Flowdeck.Mobile.Mvvm;
using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile.ViewModels;

/// <summary>A column heading in the month grid.</summary>
public sealed class WeekdayHeader
{
    public WeekdayHeader(DayOfWeek day, bool colorWeekends)
    {
        Label = day switch
        {
            DayOfWeek.Monday => "월",
            DayOfWeek.Tuesday => "화",
            DayOfWeek.Wednesday => "수",
            DayOfWeek.Thursday => "목",
            DayOfWeek.Friday => "금",
            DayOfWeek.Saturday => "토",
            _ => "일",
        };

        IsSaturday = colorWeekends && day == DayOfWeek.Saturday;
        IsSunday = colorWeekends && day == DayOfWeek.Sunday;
    }

    public string Label { get; }

    public bool IsSaturday { get; }

    public bool IsSunday { get; }
}

/// <summary>One cell in the month grid.</summary>
public sealed class CalendarDay : ObservableObject
{
    private bool _isSelected;
    private bool _hasEvents;
    private bool _hasTodos;

    public CalendarDay(DateTime date, DateTime visibleMonth, DateTime today, bool colorWeekends, Action<DateTime> onSelect)
    {
        Date = date;
        IsCurrentMonth = date.Month == visibleMonth.Month && date.Year == visibleMonth.Year;
        IsToday = date.Date == today.Date;
        IsSaturday = colorWeekends && date.DayOfWeek == DayOfWeek.Saturday;
        IsSunday = colorWeekends && date.DayOfWeek == DayOfWeek.Sunday;
        SelectCommand = new RelayCommand(() => onSelect(Date));
    }

    public DateTime Date { get; }

    public string DayNumber => Date.Day.ToString(CultureInfo.InvariantCulture);

    public bool IsCurrentMonth { get; }

    public bool IsToday { get; }

    public bool IsSaturday { get; }

    public bool IsSunday { get; }

    /// <summary>Drawn as a small dot under the number.</summary>
    public bool HasEvents
    {
        get => _hasEvents;
        set => Set(ref _hasEvents, value);
    }

    public bool HasTodos
    {
        get => _hasTodos;
        set => Set(ref _hasTodos, value);
    }

    public bool IsSelected
    {
        get => _isSelected;
        set => Set(ref _isSelected, value);
    }

    public ICommand SelectCommand { get; }
}

/// <summary>
/// The month grid, with the chosen day's entries underneath — the desktop widget's shape,
/// which is the one the calendar is actually read in.
///
/// Always six rows. A month that only needs five would otherwise make the list below jump
/// up and down as the months are paged through.
/// </summary>
public sealed class CalendarViewModel : ObservableObject
{
    private const int WeeksShown = 6;

    private readonly Func<DateTime> _clock;

    private DateTime _visibleMonth;
    private DateTime _selectedDate;

    public CalendarViewModel(Func<DateTime>? clock = null)
    {
        _clock = clock ?? (() => DateTime.Now);

        var today = _clock().Date;
        _visibleMonth = new DateTime(today.Year, today.Month, 1);
        _selectedDate = today;

        PreviousMonthCommand = new RelayCommand(() => ShiftMonth(-1));
        NextMonthCommand = new RelayCommand(() => ShiftMonth(1));
        TodayCommand = new RelayCommand(GoToToday);

        Workspace.Repository.Changed += (_, _) => MainThread.BeginInvokeOnMainThread(Refresh);

        Refresh();
    }

    /// <summary>Raised when an entry under the grid is tapped.</summary>
    public event EventHandler<(string Id, bool IsTodo)>? EditRequested;

    public ObservableCollection<WeekdayHeader> WeekdayHeaders { get; } = new();

    public ObservableCollection<CalendarDay> Days { get; } = new();

    public ObservableCollection<EventRow> Events { get; } = new();

    public ObservableCollection<TodoRow> Todos { get; } = new();

    public ICommand PreviousMonthCommand { get; }

    public ICommand NextMonthCommand { get; }

    public ICommand TodayCommand { get; }

    public string MonthLabel => _visibleMonth.ToString("yyyy년 M월", Ko.Culture);

    public string SelectedLabel => _selectedDate.ToString("M월 d일 (ddd)", Ko.Culture);

    public bool HasEvents => Events.Count > 0;

    public bool HasTodos => Todos.Count > 0;

    public bool IsEmpty => Events.Count == 0 && Todos.Count == 0;

    /// <summary>Rebuilds the grid and the day beneath it. Cheap enough to do on any change.</summary>
    public void Refresh()
    {
        var today = _clock().Date;
        var colorWeekends = Workspace.Settings.ColorWeekends;
        var firstDay = Workspace.Settings.FirstDayOfWeek;

        BuildWeekdayHeaders(firstDay, colorWeekends);
        BuildDays(today, firstDay, colorWeekends);
        LoadSelectedDay(today);

        Raise(nameof(MonthLabel));
        Raise(nameof(SelectedLabel));
    }

    public void Select(DateTime day)
    {
        _selectedDate = day.Date;

        // Tapping a trailing or leading cell should take the month with it, or the chosen
        // day would sit outside the grid that is on screen.
        if (day.Month != _visibleMonth.Month || day.Year != _visibleMonth.Year)
        {
            _visibleMonth = new DateTime(day.Year, day.Month, 1);
            Refresh();
            return;
        }

        foreach (var cell in Days) cell.IsSelected = cell.Date == _selectedDate;

        LoadSelectedDay(_clock().Date);
        Raise(nameof(SelectedLabel));
    }

    private void ShiftMonth(int months)
    {
        _visibleMonth = _visibleMonth.AddMonths(months);
        Refresh();
    }

    private void GoToToday()
    {
        var today = _clock().Date;
        _visibleMonth = new DateTime(today.Year, today.Month, 1);
        _selectedDate = today;
        Refresh();
    }

    private void BuildWeekdayHeaders(DayOfWeek firstDay, bool colorWeekends)
    {
        WeekdayHeaders.Clear();
        for (var i = 0; i < 7; i++)
        {
            WeekdayHeaders.Add(new WeekdayHeader((DayOfWeek)(((int)firstDay + i) % 7), colorWeekends));
        }
    }

    private void BuildDays(DateTime today, DayOfWeek firstDay, bool colorWeekends)
    {
        var lead = ((int)_visibleMonth.DayOfWeek - (int)firstDay + 7) % 7;
        var start = _visibleMonth.AddDays(-lead);
        var end = start.AddDays(WeeksShown * 7 - 1);

        var withEvents = Workspace.Repository.DaysWithEvents(start, end);
        var withTodos = Workspace.Repository.OpenTodos()
            .Where(t => t.DueAt.HasValue)
            .Select(t => t.DueAt!.Value.Date)
            .ToHashSet();

        Days.Clear();
        for (var i = 0; i < WeeksShown * 7; i++)
        {
            var date = start.AddDays(i);
            Days.Add(new CalendarDay(date, _visibleMonth, today, colorWeekends, Select)
            {
                HasEvents = withEvents.Contains(date),
                HasTodos = withTodos.Contains(date),
                IsSelected = date == _selectedDate,
            });
        }
    }

    private void LoadSelectedDay(DateTime today)
    {
        var now = _clock();

        Events.Clear();
        foreach (var occurrence in Workspace.Repository.OccurrencesOn(_selectedDate))
        {
            Events.Add(new EventRow(occurrence, DeleteEventAsync, EditEvent));
        }

        Todos.Clear();

        // On today, anything still open and overdue comes along. An unfinished todo that
        // vanished the day after it was due is a todo that gets forgotten.
        var todos = _selectedDate == today.Date
            ? Workspace.Repository.TodosForToday(now, includeUndated: false)
            : Workspace.Repository.TodosOn(_selectedDate);

        foreach (var todo in todos)
        {
            Todos.Add(new TodoRow(todo, now, ToggleTodoAsync, DeleteTodoAsync, EditTodo));
        }

        Raise(nameof(HasEvents));
        Raise(nameof(HasTodos));
        Raise(nameof(IsEmpty));
    }

    private void EditTodo(TodoRow row) => EditRequested?.Invoke(this, (row.Id, true));

    private void EditEvent(EventRow row) => EditRequested?.Invoke(this, (row.Id, false));

    private Task ToggleTodoAsync(TodoRow row) => Workspace.Repository.ToggleTodoAsync(row.Id, _clock());

    private Task DeleteTodoAsync(TodoRow row) => Workspace.Repository.DeleteTodoAsync(row.Id);

    private Task DeleteEventAsync(EventRow row) => Workspace.Repository.DeleteEventAsync(row.Id);
}

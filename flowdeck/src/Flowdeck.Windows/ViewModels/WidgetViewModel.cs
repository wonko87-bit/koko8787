using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Threading;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Settings;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>
/// Drives the desktop widget: a month grid on top, the selected day's events and
/// todos below, and an inline capture box at the bottom that speaks the same
/// natural-language syntax as the pop-up window.
/// </summary>
public sealed class WidgetViewModel : ObservableObject
{
    private readonly WorkspaceRepository _repository;
    private readonly NaturalLanguageParser _parser;
    private readonly Func<DateTime> _clock;
    private readonly Dispatcher? _dispatcher;

    private DateTime _visibleMonth;
    private DateTime _selectedDate;
    private DateTime _knownToday;
    private string _quickInput = string.Empty;
    private string _previewText = string.Empty;
    private bool _hasPreview;
    private double _dayCellSize = 38;
    private double _dayFontSize = 12;

    private DayOfWeek _firstDayOfWeek = DayOfWeek.Monday;
    private bool _colorWeekends;

    public WidgetViewModel(
        WorkspaceRepository repository,
        NaturalLanguageParser parser,
        Func<DateTime>? clock = null)
    {
        _repository = repository;
        _parser = parser;
        _clock = clock ?? (() => DateTime.Now);

        var today = _clock().Date;
        _visibleMonth = new DateTime(today.Year, today.Month, 1);
        _selectedDate = today;
        _knownToday = today;

        PreviousMonthCommand = new RelayCommand(() => ShiftMonth(-1));
        NextMonthCommand = new RelayCommand(() => ShiftMonth(1));
        TodayCommand = new RelayCommand(GoToToday);
        SubmitCommand = new AsyncRelayCommand(SubmitAsync, () => !string.IsNullOrWhiteSpace(QuickInput));

        // The view model is built on the UI thread; a repository change raised anywhere
        // else still has to land here, because these collections are bound to the widget.
        _dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
        _repository.Changed += (_, _) => OnRepositoryChanged();

        BuildWeekdayHeaders();
        Refresh();
    }

    /// <summary>Raised once a line typed into the widget's box has been filed.</summary>
    public event EventHandler<CaptureResult>? Captured;

    public ObservableCollection<CalendarDay> Days { get; } = new();

    public ObservableCollection<EventRow> Events { get; } = new();

    public ObservableCollection<TodoRow> Todos { get; } = new();

    /// <summary>Column headers, starting on whichever day the user chose.</summary>
    public ObservableCollection<WeekdayHeader> WeekdayHeaders { get; } = new();

    /// <summary>
    /// Edge length of one day cell. Driven by the widget's width so the grid keeps its
    /// proportions as the window grows, which is also what moves the divider below it.
    /// </summary>
    public double DayCellSize
    {
        get => _dayCellSize;
        private set => Set(ref _dayCellSize, value);
    }

    /// <summary>Scales with the cell so a larger grid does not end up mostly empty space.</summary>
    public double DayFontSize
    {
        get => _dayFontSize;
        private set => Set(ref _dayFontSize, value);
    }

    public ICommand PreviousMonthCommand { get; }

    public ICommand NextMonthCommand { get; }

    public ICommand TodayCommand { get; }

    public AsyncRelayCommand SubmitCommand { get; }

    public string MonthTitle => _visibleMonth.ToString("yyyy년 M월", Ko.Culture);

    public string SelectedDateLabel => _selectedDate.ToString("M월 d일 dddd", Ko.Culture);

    public bool SelectedIsToday => _selectedDate.Date == _clock().Date;

    public string TodoSectionLabel => SelectedIsToday ? "할일" : "할일 · " + _selectedDate.ToString("M월 d일", Ko.Culture);

    public bool HasEvents => Events.Count > 0;

    public bool HasTodos => Todos.Count > 0;

    public string QuickInput
    {
        get => _quickInput;
        set
        {
            if (!Set(ref _quickInput, value)) return;
            UpdatePreview();
            SubmitCommand.RaiseCanExecuteChanged();
        }
    }

    /// <summary>A live reading of the input box: where it will go and when.</summary>
    public string PreviewText
    {
        get => _previewText;
        private set => Set(ref _previewText, value);
    }

    public bool HasPreview
    {
        get => _hasPreview;
        private set => Set(ref _hasPreview, value);
    }

    public DateTime SelectedDate
    {
        get => _selectedDate;
        private set
        {
            if (!Set(ref _selectedDate, value)) return;
            Raise(nameof(SelectedDateLabel));
            Raise(nameof(SelectedIsToday));
            Raise(nameof(TodoSectionLabel));
        }
    }

    public void Refresh()
    {
        BuildMonth();
        BuildDayLists();
    }

    /// <summary>Re-reads the calendar options after the settings window changes them.</summary>
    public void ApplyCalendarSettings(AppSettings settings)
    {
        var changed = _firstDayOfWeek != settings.FirstDayOfWeek || _colorWeekends != settings.ColorWeekends;
        if (!changed) return;

        _firstDayOfWeek = settings.FirstDayOfWeek;
        _colorWeekends = settings.ColorWeekends;

        BuildWeekdayHeaders();
        Refresh();
    }

    /// <summary>
    /// Recomputes the grid metrics for the widget's current content width. Cells are square,
    /// so the calendar's height follows from the width and the layout stays in proportion.
    /// </summary>
    public void UpdateMetrics(double contentWidth)
    {
        if (contentWidth <= 0) return;

        DayCellSize = Math.Clamp(contentWidth / 7d, MinDayCell, MaxDayCell);
        DayFontSize = Math.Clamp(DayCellSize * 0.30, 11d, 16d);
    }

    private const double MinDayCell = 30;
    private const double MaxDayCell = 62;

    private void BuildWeekdayHeaders()
    {
        WeekdayHeaders.Clear();
        for (var i = 0; i < 7; i++)
        {
            var day = (DayOfWeek)(((int)_firstDayOfWeek + i) % 7);
            WeekdayHeaders.Add(new WeekdayHeader(day, _colorWeekends));
        }
    }

    /// <summary>
    /// Called on a timer. A widget that sits on the desktop for days has to notice
    /// midnight passing, otherwise "오늘" keeps pointing at yesterday.
    /// </summary>
    public void TickClock()
    {
        var today = _clock().Date;
        if (today == _knownToday) return;

        var wasFollowingToday = _selectedDate.Date == _knownToday;
        _knownToday = today;

        if (wasFollowingToday)
        {
            GoToToday();
        }
        else
        {
            Refresh();
        }
    }

    private void OnRepositoryChanged()
    {
        if (_dispatcher is null || _dispatcher.CheckAccess())
        {
            Refresh();
            return;
        }

        _dispatcher.Invoke(Refresh);
    }

    private void ShiftMonth(int delta)
    {
        _visibleMonth = _visibleMonth.AddMonths(delta);
        Raise(nameof(MonthTitle));
        BuildMonth();
    }

    private void GoToToday()
    {
        var today = _clock().Date;
        _visibleMonth = new DateTime(today.Year, today.Month, 1);
        Raise(nameof(MonthTitle));
        SelectDay(today);
    }

    private void SelectDay(DateTime date)
    {
        SelectedDate = date.Date;

        if (date.Month != _visibleMonth.Month || date.Year != _visibleMonth.Year)
        {
            _visibleMonth = new DateTime(date.Year, date.Month, 1);
            Raise(nameof(MonthTitle));
        }

        BuildMonth();
        BuildDayLists();
    }

    private void BuildMonth()
    {
        var today = _clock().Date;

        // Six full weeks starting on the week-opening day on or before the 1st, so the
        // grid never changes height as the user pages through months.
        var gridStart = _visibleMonth.AddDays(-WeekIndex(_visibleMonth.DayOfWeek));
        var gridEnd = gridStart.AddDays(41);

        var eventDays = _repository.DaysWithEvents(gridStart, gridEnd);
        var todoDays = _repository.OpenTodos()
            .Where(t => t.DueAt.HasValue)
            .Select(t => t.DueAt!.Value.Date)
            .ToHashSet();

        Days.Clear();
        for (var i = 0; i < 42; i++)
        {
            var date = gridStart.AddDays(i);
            Days.Add(new CalendarDay(date, _visibleMonth, today, _colorWeekends, SelectDay)
            {
                HasEvents = eventDays.Contains(date),
                HasTodos = todoDays.Contains(date),
                IsSelected = date == _selectedDate.Date,
            });
        }
    }

    private void BuildDayLists()
    {
        var now = _clock();

        Events.Clear();
        foreach (var occurrence in _repository.OccurrencesOn(_selectedDate))
        {
            Events.Add(new EventRow(occurrence, DeleteEventAsync));
        }

        var todos = SelectedIsToday
            ? _repository.TodosForToday(now)
            : _repository.TodosOn(_selectedDate);

        Todos.Clear();
        foreach (var todo in todos)
        {
            Todos.Add(new TodoRow(todo, now, ToggleTodoAsync, DeleteTodoAsync));
        }

        Raise(nameof(HasEvents));
        Raise(nameof(HasTodos));
    }

    private void UpdatePreview()
    {
        if (string.IsNullOrWhiteSpace(_quickInput))
        {
            HasPreview = false;
            PreviewText = string.Empty;
            return;
        }

        var parsed = _parser.Parse(_quickInput, _clock());
        HasPreview = true;
        PreviewText = $"{parsed.DescribeTarget()} · {parsed.DescribeSchedule()}";
    }

    private async Task SubmitAsync()
    {
        var text = _quickInput;
        if (string.IsNullOrWhiteSpace(text)) return;

        var now = _clock();
        var parsed = _parser.Parse(text, now);
        var result = await _repository.CaptureAsync(parsed, now);

        QuickInput = string.Empty;

        // The widget's own box files entries just as the overlay does, so whatever the
        // shell does afterwards — reporting a failed push, opening the Teams form — has
        // to happen here too, or the same line would behave differently in two places.
        if (!result.IsEmpty) Captured?.Invoke(this, result);

        // Follow the entry so the user sees where it landed.
        if (parsed.Start.HasValue) SelectDay(parsed.Start.Value.Date);
    }

    private Task ToggleTodoAsync(TodoRow row) => _repository.ToggleTodoAsync(row.Id, _clock());

    private Task DeleteTodoAsync(TodoRow row) => _repository.DeleteTodoAsync(row.Id);

    private Task DeleteEventAsync(EventRow row) => _repository.DeleteEventAsync(row.Id);

    private int WeekIndex(DayOfWeek day) => ((int)day - (int)_firstDayOfWeek + 7) % 7;
}

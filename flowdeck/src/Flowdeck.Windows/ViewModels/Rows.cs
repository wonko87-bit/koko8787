using System.Globalization;
using System.Windows.Input;
using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Flowdeck.Windows.Infrastructure;
using Flowdeck.Windows.Services;

namespace Flowdeck.Windows.ViewModels;

internal static class Ko
{
    public static readonly CultureInfo Culture = CultureInfo.GetCultureInfo("ko-KR");
}

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

        IsSaturday = day == DayOfWeek.Saturday;
        IsSunday = day == DayOfWeek.Sunday;
        ColorWeekends = colorWeekends;
    }

    public string Label { get; }

    public bool IsSaturday { get; }

    public bool IsSunday { get; }

    public bool ColorWeekends { get; }
}

/// <summary>One cell in the month grid.</summary>
public sealed class CalendarDay : ObservableObject
{
    private bool _isSelected;

    public CalendarDay(
        DateTime date,
        DateTime visibleMonth,
        DateTime today,
        bool colorWeekends,
        Action<DateTime> onSelect)
    {
        Date = date;
        IsCurrentMonth = date.Month == visibleMonth.Month && date.Year == visibleMonth.Year;
        IsToday = date.Date == today.Date;
        IsSaturday = date.DayOfWeek == DayOfWeek.Saturday;
        IsSunday = date.DayOfWeek == DayOfWeek.Sunday;
        IsWeekend = IsSaturday || IsSunday;
        ColorWeekends = colorWeekends;
        SelectCommand = new RelayCommand(() => onSelect(Date));
    }

    public DateTime Date { get; }

    public string DayNumber => Date.Day.ToString(CultureInfo.InvariantCulture);

    public bool IsCurrentMonth { get; }

    public bool IsToday { get; }

    public bool IsWeekend { get; }

    public bool IsSaturday { get; }

    public bool IsSunday { get; }

    /// <summary>Carried per cell so the template can decide without reaching for global state.</summary>
    public bool ColorWeekends { get; }

    /// <summary>Drawn as a small dot under the number.</summary>
    public bool HasEvents { get; set; }

    public bool HasTodos { get; set; }

    public bool IsSelected
    {
        get => _isSelected;
        set => Set(ref _isSelected, value);
    }

    public ICommand SelectCommand { get; }
}

/// <summary>
/// A meeting read out of Outlook that Flowdeck did not put there — someone else's booking,
/// a room, a leave day. Read-only: nothing here may be ticked, edited or deleted. The one
/// thing it can do is be taken in — a double-click makes a local copy to keep a note on,
/// and from then on the copy is the row that shows.
/// </summary>
public sealed class ExternalRow
{
    public ExternalRow(ExternalOccurrence occurrence, Func<ExternalRow, Task>? adopt = null)
    {
        Occurrence = occurrence;
        AdoptCommand = new AsyncRelayCommand(() => adopt?.Invoke(this) ?? Task.CompletedTask);

        Title = string.IsNullOrWhiteSpace(occurrence.Title) ? "(제목 없음)" : occurrence.Title;
        Location = occurrence.Location;
        HasLocation = !string.IsNullOrWhiteSpace(occurrence.Location);

        TimeLabel = occurrence.IsAllDay
            ? "종일"
            : occurrence.Start.ToString("tt h:mm", Ko.Culture);

        RangeLabel = occurrence.IsAllDay
            ? string.Empty
            : $"{occurrence.Start:HH:mm} – {occurrence.End:HH:mm}";
    }

    public string Title { get; }

    public string TimeLabel { get; }

    public string RangeLabel { get; }

    public string Location { get; }

    public bool HasLocation { get; }

    public ExternalOccurrence Occurrence { get; }

    /// <summary>Takes a copy of the meeting and opens it. Bound to a double-click, like the editable rows.</summary>
    public ICommand AdoptCommand { get; }
}

/// <summary>A todo in the widget list.</summary>
public sealed class TodoRow : ObservableObject
{
    private readonly TodoItem _item;
    private readonly Func<TodoRow, Task> _toggle;

    public TodoRow(
        TodoItem item,
        DateTime now,
        Func<TodoRow, Task> toggle,
        Func<TodoRow, Task> delete,
        Action<TodoRow>? edit = null)
    {
        _item = item;
        _toggle = toggle;
        IsOverdue = item.IsOverdue(now);
        DeleteCommand = new AsyncRelayCommand(() => delete(this));
        EditCommand = new RelayCommand(() => edit?.Invoke(this));
    }

    public string Id => _item.Id;

    public string Title => _item.Title;

    public bool IsOverdue { get; }

    public bool HasTags => _item.Tags.Count > 0;

    public string TagLabel => string.Join("  ", _item.Tags.Select(t => "#" + t));

    public bool IsRecurring => _item.Recurrence.IsRepeating;

    public string RecurrenceLabel => _item.Recurrence.Describe();

    /// <summary>Priority is shown as a bar of ticks, never as a colour.</summary>
    public string PriorityLabel => _item.Priority switch
    {
        Priority.Urgent => "!!!",
        Priority.High => "!!",
        Priority.Normal => "!",
        _ => string.Empty,
    };

    public bool HasPriority => _item.Priority >= Priority.Normal;

    /// <summary>The first line of the note, so a row says whether there is more to read.</summary>
    public string NotePreview => FirstLine(_item.Notes);

    /// <summary>The whole note, for the tooltip — the only place a long one fits.</summary>
    public string Notes => _item.Notes;

    public bool HasNote => _item.Notes.Length > 0;

    internal static string FirstLine(string notes)
    {
        if (string.IsNullOrEmpty(notes)) return string.Empty;

        var end = notes.IndexOf('\n');
        var first = end < 0 ? notes : notes[..end];
        return end < 0 ? first : first + " …";
    }

    public string DueLabel
    {
        get
        {
            if (!_item.DueAt.HasValue) return string.Empty;
            var due = _item.DueAt.Value;

            if (!_item.HasTime) return due.ToString("M월 d일", Ko.Culture);

            // An overdue item carries its date, because it shows up under today: a bare
            // "오후 9:00" there reads as due tonight rather than as left over from before.
            return IsOverdue
                ? due.ToString("M월 d일 tt h:mm", Ko.Culture)
                : due.ToString("tt h:mm", Ko.Culture);
        }
    }

    public bool HasDue => _item.DueAt.HasValue;

    /// <summary>
    /// Two-way bound to the checkbox. Setting it hands the work to the repository rather
    /// than mutating the model here, so a repeating todo can advance instead of closing.
    /// </summary>
    public bool IsDone
    {
        get => _item.IsDone;
        set
        {
            if (value == _item.IsDone) return;

            // The checkbox has no way to await this, so the task is watched instead of
            // being dropped on the floor.
            CrashReporter.Observe(_toggle(this), "TodoRow.IsDone");
            Raise();
        }
    }

    public ICommand DeleteCommand { get; }

    /// <summary>Opens the detail window. Bound to a double-click, the desktop way to open a row.</summary>
    public ICommand EditCommand { get; }
}

/// <summary>
/// A repeat rule collapsed to one line — "3분할 운동 - 상체 [매주 월,목]".
///
/// The list windows show these instead of scattering every occurrence through the date
/// groups, which would bury the one-off entries under a rule that never ends.
/// </summary>
public sealed class RecurringRow : ObservableObject
{
    private readonly TodoItem? _todo;
    private readonly Func<RecurringRow, Task>? _toggle;

    /// <summary>A repeating todo. Its due date is the occurrence currently pending.</summary>
    public RecurringRow(TodoItem todo, Func<RecurringRow, Task> toggle, Func<RecurringRow, Task> delete)
    {
        _todo = todo;
        _toggle = toggle;

        Id = todo.Id;
        Title = todo.Title;
        IsTodo = true;
        PatternLabel = Wrap(todo.Recurrence.DescribeShort(todo.DueAt));
        NextLabel = todo.DueAt.HasValue ? "다음 " + FormatNext(todo.DueAt.Value, todo.HasTime) : string.Empty;
        TagLabel = string.Join("  ", todo.Tags.Select(t => "#" + t));
        HasTags = todo.Tags.Count > 0;
        DeleteCommand = new AsyncRelayCommand(() => delete(this));
    }

    /// <summary>A repeating event. The next occurrence is worked out from the rule.</summary>
    public RecurringRow(CalendarEvent source, DateTime now, Func<RecurringRow, Task> delete)
    {
        Id = source.Id;
        Title = source.Title;
        IsTodo = false;
        PatternLabel = Wrap(source.Recurrence.DescribeShort(source.Start));

        var next = RecurrenceExpander.Next(source.Start, source.Recurrence, now);
        NextLabel = next.HasValue ? "다음 " + FormatNext(next.Value, !source.IsAllDay) : string.Empty;

        TagLabel = string.Join("  ", source.Tags.Select(t => "#" + t));
        HasTags = source.Tags.Count > 0;
        DeleteCommand = new AsyncRelayCommand(() => delete(this));
    }

    public string Id { get; }

    public string Title { get; }

    /// <summary>The repeat rule in brackets, e.g. "[매주 월,목]".</summary>
    public string PatternLabel { get; }

    public string NextLabel { get; }

    public string TagLabel { get; }

    public bool HasTags { get; }

    /// <summary>Only todos get a checkbox; an event has nothing to complete.</summary>
    public bool IsTodo { get; }

    /// <summary>Removes the repeat rule itself, not one occurrence of it.</summary>
    public ICommand DeleteCommand { get; }

    /// <summary>
    /// Ticking a repeating todo does not close it — the repository rolls its due date on
    /// to the next occurrence, so this row stays put and its "다음" date moves.
    /// </summary>
    public bool IsDone
    {
        get => _todo?.IsDone ?? false;
        set
        {
            if (_todo is null || _toggle is null || value == _todo.IsDone) return;

            CrashReporter.Observe(_toggle(this), "RecurringRow.IsDone");
            Raise();
        }
    }

    private static string Wrap(string pattern) => pattern.Length == 0 ? string.Empty : $"[{pattern}]";

    private static string FormatNext(DateTime when, bool hasTime) =>
        hasTime ? when.ToString("M월 d일 HH:mm", Ko.Culture) : when.ToString("M월 d일", Ko.Culture);
}

/// <summary>One occurrence of an event in the day list.</summary>
public sealed class EventRow : ObservableObject
{
    public EventRow(EventOccurrence occurrence, Func<EventRow, Task> delete, Action<EventRow>? edit = null)
    {
        EditCommand = new RelayCommand(() => edit?.Invoke(this));
        Id = occurrence.Source.Id;
        Notes = occurrence.Source.Notes;
        NotePreview = TodoRow.FirstLine(occurrence.Source.Notes);
        HasNote = occurrence.Source.Notes.Length > 0;
        Title = occurrence.Title;
        IsAllDay = occurrence.IsAllDay;
        IsRecurring = occurrence.IsRecurring;
        RecurrenceLabel = occurrence.Source.Recurrence.Describe();
        TagLabel = string.Join("  ", occurrence.Source.Tags.Select(t => "#" + t));
        HasTags = occurrence.Source.Tags.Count > 0;

        TimeLabel = occurrence.IsAllDay
            ? "종일"
            : occurrence.Start.ToString("HH:mm", Ko.Culture);

        RangeLabel = occurrence.IsAllDay
            ? string.Empty
            : $"{occurrence.Start:HH:mm} – {occurrence.End:HH:mm}";

        DeleteCommand = new AsyncRelayCommand(() => delete(this));
    }

    public string Id { get; }

    public string Title { get; }

    public string TimeLabel { get; }

    public string RangeLabel { get; }

    public bool IsAllDay { get; }

    public bool IsRecurring { get; }

    public string RecurrenceLabel { get; }

    public string TagLabel { get; }

    public bool HasTags { get; }

    public string NotePreview { get; }

    public string Notes { get; }

    public bool HasNote { get; }

    public ICommand DeleteCommand { get; }

    /// <summary>Opens the detail window. Bound to a double-click, the desktop way to open a row.</summary>
    public ICommand EditCommand { get; }
}

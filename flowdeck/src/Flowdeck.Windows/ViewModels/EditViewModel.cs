using System.Globalization;
using System.Windows.Input;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>
/// One entry, opened up and editable — the desktop twin of the mobile detail sheet.
///
/// Before this, the only way to change something on the desktop was to delete it and type
/// it again. That is bearable for a typo in a title and hopeless for a note, which is the
/// whole reason the window exists.
/// </summary>
public sealed class EditViewModel : ObservableObject
{
    private readonly WorkspaceRepository _repository;
    private readonly Func<DateTime> _clock;
    private readonly TodoItem? _todo;
    private readonly CalendarEvent? _event;

    private string _title;
    private string _notes;
    private bool _hasDate;
    private bool _hasTime;
    private string _date;
    private string _time;
    private string _tags;
    private Priority _priority;
    private string _reminder;
    private string _status = string.Empty;

    private EditViewModel(
        WorkspaceRepository repository,
        TodoItem? todo,
        CalendarEvent? source,
        Func<DateTime>? clock)
    {
        _repository = repository;
        _clock = clock ?? (() => DateTime.Now);
        _todo = todo;
        _event = source;

        var now = _clock();
        var when = todo?.DueAt ?? source?.Start;

        _title = todo?.Title ?? source?.Title ?? string.Empty;
        _notes = todo?.Notes ?? source?.Notes ?? string.Empty;
        _hasDate = when.HasValue;
        _hasTime = todo?.HasTime ?? (source is not null && !source.IsAllDay);
        _date = (when ?? now).ToString("yyyy-MM-dd", Ko.Culture);
        _time = (when ?? now).ToString("HH:mm", Ko.Culture);
        _tags = string.Join(" ", (todo?.Tags ?? source?.Tags ?? new List<string>()).Select(t => "#" + t));
        _priority = todo?.Priority ?? Priority.None;

        var minutes = todo?.ReminderMinutesBefore ?? source?.ReminderMinutesBefore;
        _reminder = minutes?.ToString(Ko.Culture) ?? string.Empty;

        SaveCommand = new AsyncRelayCommand(SaveAsync);
        DeleteCommand = new AsyncRelayCommand(DeleteAsync);
    }

    /// <summary>
    /// Null when the entry has gone — deleted from another window between the double-click
    /// and this call, which two list windows over the same repository make possible.
    /// </summary>
    public static EditViewModel? For(
        WorkspaceRepository repository,
        string id,
        bool isTodo,
        Func<DateTime>? clock = null)
    {
        if (isTodo)
        {
            var todo = repository.Todos.FirstOrDefault(t => t.Id == id);
            return todo is null ? null : new EditViewModel(repository, todo, null, clock);
        }

        var source = repository.Events.FirstOrDefault(e => e.Id == id);
        return source is null ? null : new EditViewModel(repository, null, source, clock);
    }

    /// <summary>Raised once the entry has been saved or deleted, so the window can close.</summary>
    public event EventHandler? Finished;

    public bool IsTodo => _todo is not null;

    public string HeaderLabel => IsTodo ? "할일 편집" : "일정 편집";

    /// <summary>Priority belongs to todos; an event has no such field to edit.</summary>
    public bool ShowPriority => IsTodo;

    /// <summary>An event has to sit somewhere on the calendar, so only a todo may lose its date.</summary>
    public bool CanClearDate => IsTodo;

    public ICommand SaveCommand { get; }

    public ICommand DeleteCommand { get; }

    public string Title
    {
        get => _title;
        set => Set(ref _title, value);
    }

    public string Notes
    {
        get => _notes;
        set => Set(ref _notes, value);
    }

    public bool HasDate
    {
        get => _hasDate;
        set
        {
            if (!Set(ref _hasDate, value)) return;

            if (!value) HasTime = false;
            Raise(nameof(ShowTime));
        }
    }

    public bool HasTime
    {
        get => _hasTime;
        set
        {
            if (Set(ref _hasTime, value)) Raise(nameof(ShowTime));
        }
    }

    public bool ShowTime => _hasDate && _hasTime;

    /// <summary>Typed as text rather than picked, because the keyboard is already under the hands.</summary>
    public string Date
    {
        get => _date;
        set
        {
            if (Set(ref _date, value)) Status = string.Empty;
        }
    }

    public string Time
    {
        get => _time;
        set
        {
            if (Set(ref _time, value)) Status = string.Empty;
        }
    }

    /// <summary>Typed with or without the hashes; either way the same tags come out.</summary>
    public string Tags
    {
        get => _tags;
        set => Set(ref _tags, value);
    }

    public string Reminder
    {
        get => _reminder;
        set => Set(ref _reminder, value);
    }

    public bool IsUrgent
    {
        get => _priority == Priority.Urgent;
        set => Choose(Priority.Urgent, value);
    }

    public bool IsHigh
    {
        get => _priority == Priority.High;
        set => Choose(Priority.High, value);
    }

    public bool IsNormal
    {
        get => _priority == Priority.Normal;
        set => Choose(Priority.Normal, value);
    }

    public bool IsLow
    {
        get => _priority == Priority.Low;
        set => Choose(Priority.Low, value);
    }

    public bool IsNoPriority
    {
        get => _priority == Priority.None;
        set => Choose(Priority.None, value);
    }

    public string Status
    {
        get => _status;
        set
        {
            if (Set(ref _status, value)) Raise(nameof(HasStatus));
        }
    }

    public bool HasStatus => _status.Length > 0;

    public async Task SaveAsync()
    {
        if (!TryBuildWhen(out var when))
        {
            Status = "날짜는 2026-08-12, 시각은 14:30 처럼 써 주세요";
            return;
        }

        var edit = new EntryEdit
        {
            Title = Title,
            Notes = Notes ?? string.Empty,
            When = when,
            HasTime = _hasTime,
            Priority = _priority,
            Tags = SplitTags(_tags),
            ReminderMinutesBefore =
                int.TryParse(_reminder, out var minutes) && minutes > 0 ? minutes : null,
        };

        var saved = IsTodo
            ? await _repository.UpdateTodoAsync(_todo!.Id, edit, _clock())
            : await _repository.UpdateEventAsync(_event!.Id, edit, _clock());

        if (!saved)
        {
            Status = "이미 삭제된 항목입니다";
            return;
        }

        Finished?.Invoke(this, EventArgs.Empty);
    }

    public async Task DeleteAsync()
    {
        if (IsTodo) await _repository.DeleteTodoAsync(_todo!.Id);
        else await _repository.DeleteEventAsync(_event!.Id);

        Finished?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Reads the two text boxes back into one instant. Refuses rather than guesses: a date
    /// that will not parse means saving would silently move the entry somewhere else.
    /// </summary>
    private bool TryBuildWhen(out DateTime? when)
    {
        when = null;
        if (!_hasDate) return true;

        if (!DateTime.TryParse(_date, Ko.Culture, DateTimeStyles.None, out var day)) return false;

        if (!_hasTime)
        {
            when = day.Date;
            return true;
        }

        if (!TimeSpan.TryParse(_time, Ko.Culture, out var clock)) return false;
        if (clock < TimeSpan.Zero || clock >= TimeSpan.FromDays(1)) return false;

        when = day.Date + clock;
        return true;
    }

    /// <summary>
    /// The levels are radio buttons, so only the one being switched on has anything to say.
    /// Every level is raised afterwards because turning one on turns another off.
    /// </summary>
    private void Choose(Priority level, bool on)
    {
        if (!on || _priority == level) return;

        _priority = level;
        Raise(nameof(IsUrgent));
        Raise(nameof(IsHigh));
        Raise(nameof(IsNormal));
        Raise(nameof(IsLow));
        Raise(nameof(IsNoPriority));
    }

    /// <summary>
    /// Splits on whitespace and on the hashes themselves, so "#업무#보고", "#업무 #보고" and
    /// "업무 보고" all give the same two tags.
    /// </summary>
    internal static List<string> SplitTags(string? text) =>
        (text ?? string.Empty)
            .Split(new[] { ' ', '\t', '\n', ',', '#' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(t => t.Trim())
            .Where(t => t.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
}

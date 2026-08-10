using System.Windows.Input;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Flowdeck.Mobile.Mvvm;
using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile.ViewModels;

/// <summary>
/// The detail sheet: everything about one entry, editable.
///
/// Until this existed the only way to change an entry was to delete it and type it again,
/// which is fine for a typo and hopeless for a note. Notes are the reason it was built, so
/// the note field is the tall one.
/// </summary>
public sealed class EditViewModel : ObservableObject
{
    private readonly Func<DateTime> _clock;
    private readonly TodoItem? _todo;
    private readonly CalendarEvent? _event;

    private string _title;
    private string _notes;
    private bool _hasDate;
    private bool _hasTime;
    private DateTime _date;
    private TimeSpan _time;
    private string _tags;
    private Priority _priority;
    private string _reminder;
    private string _status = string.Empty;

    private EditViewModel(TodoItem? todo, CalendarEvent? source, Func<DateTime>? clock)
    {
        _clock = clock ?? (() => DateTime.Now);
        _todo = todo;
        _event = source;

        var now = _clock();
        var when = todo?.DueAt ?? source?.Start;

        _title = todo?.Title ?? source?.Title ?? string.Empty;
        _notes = todo?.Notes ?? source?.Notes ?? string.Empty;
        _hasDate = when.HasValue;
        _hasTime = todo?.HasTime ?? (source is not null && !source.IsAllDay);
        _date = (when ?? now).Date;
        _time = (when ?? now).TimeOfDay;
        _tags = string.Join(" ", (todo?.Tags ?? source?.Tags ?? new List<string>()).Select(t => "#" + t));
        _priority = todo?.Priority ?? Priority.None;

        var minutes = todo?.ReminderMinutesBefore ?? source?.ReminderMinutesBefore;
        _reminder = minutes?.ToString() ?? string.Empty;

        SaveCommand = new AsyncRelayCommand(SaveAsync);
        DeleteCommand = new AsyncRelayCommand(DeleteAsync);
        ClearDateCommand = new RelayCommand(() => HasDate = false);

        PriorityCommands = new ICommand[]
        {
            new RelayCommand(() => TogglePriority(Priority.Urgent)),
            new RelayCommand(() => TogglePriority(Priority.High)),
            new RelayCommand(() => TogglePriority(Priority.Normal)),
            new RelayCommand(() => TogglePriority(Priority.Low)),
        };
    }

    /// <summary>Null when the entry has gone — deleted from another screen while this opened.</summary>
    public static EditViewModel? For(string id, bool isTodo, Func<DateTime>? clock = null)
    {
        if (isTodo)
        {
            var todo = Workspace.Repository.Todos.FirstOrDefault(t => t.Id == id);
            return todo is null ? null : new EditViewModel(todo, null, clock);
        }

        var source = Workspace.Repository.Events.FirstOrDefault(e => e.Id == id);
        return source is null ? null : new EditViewModel(null, source, clock);
    }

    /// <summary>Raised once the entry has been saved or deleted, so the sheet can close.</summary>
    public event EventHandler? Finished;

    public bool IsTodo => _todo is not null;

    public string HeaderLabel => IsTodo ? "할일" : "일정";

    public ICommand SaveCommand { get; }

    public ICommand DeleteCommand { get; }

    public ICommand ClearDateCommand { get; }

    public IReadOnlyList<ICommand> PriorityCommands { get; }

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

    /// <summary>
    /// An event has to sit somewhere on the calendar, so only a todo may lose its date.
    /// </summary>
    public bool CanClearDate => IsTodo;

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

    public DateTime Date
    {
        get => _date;
        set => Set(ref _date, value);
    }

    public TimeSpan Time
    {
        get => _time;
        set => Set(ref _time, value);
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

    public bool IsUrgent => _priority == Priority.Urgent;

    public bool IsHigh => _priority == Priority.High;

    public bool IsNormal => _priority == Priority.Normal;

    public bool IsLow => _priority == Priority.Low;

    public string Status
    {
        get => _status;
        set
        {
            if (Set(ref _status, value)) Raise(nameof(HasStatus));
        }
    }

    public bool HasStatus => _status.Length > 0;

    /// <summary>Choosing the level already set clears it, which is the only way back to none.</summary>
    private void TogglePriority(Priority level)
    {
        _priority = _priority == level ? Priority.None : level;
        Raise(nameof(IsUrgent));
        Raise(nameof(IsHigh));
        Raise(nameof(IsNormal));
        Raise(nameof(IsLow));
    }

    public async Task SaveAsync()
    {
        var edit = new EntryEdit
        {
            Title = Title,
            Notes = Notes ?? string.Empty,
            When = _hasDate ? (_hasTime ? _date.Date + _time : _date.Date) : null,
            HasTime = _hasTime,
            Priority = _priority,
            Tags = SplitTags(_tags),
            ReminderMinutesBefore = int.TryParse(_reminder, out var minutes) && minutes > 0 ? minutes : null,
        };

        var saved = IsTodo
            ? await Workspace.Repository.UpdateTodoAsync(_todo!.Id, edit, _clock())
            : await Workspace.Repository.UpdateEventAsync(_event!.Id, edit, _clock());

        if (!saved)
        {
            Status = "이미 삭제된 항목입니다";
            return;
        }

        Finished?.Invoke(this, EventArgs.Empty);
    }

    public async Task DeleteAsync()
    {
        if (IsTodo) await Workspace.Repository.DeleteTodoAsync(_todo!.Id);
        else await Workspace.Repository.DeleteEventAsync(_event!.Id);

        Finished?.Invoke(this, EventArgs.Empty);
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

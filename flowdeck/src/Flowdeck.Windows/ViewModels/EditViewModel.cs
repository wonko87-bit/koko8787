using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
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

    /// <summary>The file this entry is about, when it came from an application that named one.</summary>
    private readonly string? _file;

    private string _newAttachedTitle = string.Empty;

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

        // Read from the note as it was when the window opened. Editing the note text does not
        // move the buttons around underneath the person typing.
        _file = AttachedFile.PathIn(_notes);

        SaveCommand = new AsyncRelayCommand(SaveAsync);
        DeleteCommand = new AsyncRelayCommand(DeleteAsync);
        OpenFileCommand = new RelayCommand(OpenFile, () => CanOpenFile);
        ShowFileCommand = new RelayCommand(ShowFile, () => HasFile);
        AttachCommand = new AsyncRelayCommand(AttachAsync);
        RegisterNoteActionsCommand = new AsyncRelayCommand(RegisterNoteActionsAsync);
        DetachSelfCommand = new AsyncRelayCommand(DetachSelfAsync);

        LoadAttached();
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

    /// <summary>
    /// A meeting taken in from Outlook. Said on the window because the one thing a person
    /// will wonder, with the cursor in the note box, is whether what they type ends up in
    /// the meeting everybody else can see. It does not.
    /// </summary>
    public bool IsAdopted => _event?.Origin is not null;

    public string AdoptedNote =>
        $"{_repository.External?.DisplayName ?? "Outlook"} 에서 가져온 회의입니다. 여기 적는 메모는 이 PC에만 남고 회의 쪽으로 보내지 않습니다.";

    /// <summary>Priority belongs to todos; an event has no such field to edit.</summary>
    public bool ShowPriority => IsTodo;

    /// <summary>An event has to sit somewhere on the calendar, so only a todo may lose its date.</summary>
    public bool CanClearDate => IsTodo;

    public ICommand SaveCommand { get; }

    public ICommand DeleteCommand { get; }

    public ICommand OpenFileCommand { get; }

    public ICommand ShowFileCommand { get; }

    public ICommand AttachCommand { get; }

    public ICommand RegisterNoteActionsCommand { get; }

    public ICommand DetachSelfCommand { get; }

    // ---- todos attached to this event ----------------------------------------

    /// <summary>Only an event has things attached to it; a todo is one of those things.</summary>
    public bool ShowAttached => !IsTodo;

    public ObservableCollection<AttachedTodoRow> AttachedTodos { get; } = new();

    public bool HasAttachedTodos => AttachedTodos.Count > 0;

    /// <summary>Typed here and attached with Enter. Due when the event is; the date can be changed on the todo afterwards.</summary>
    public string NewAttachedTitle
    {
        get => _newAttachedTitle;
        set => Set(ref _newAttachedTitle, value);
    }

    /// <summary>
    /// The 할일: lines in the note that are not yet todos. Read from the box as it is now
    /// rather than as it was saved, because the moment a person wants this is the moment
    /// after typing them.
    /// </summary>
    public IReadOnlyList<string> UnregisteredNoteActions
    {
        get
        {
            if (IsTodo) return Array.Empty<string>();

            var have = new HashSet<string>(AttachedTodos.Select(t => t.Title), StringComparer.OrdinalIgnoreCase);
            return NoteOutline.Parse(Notes).Actions.Where(a => !have.Contains(a)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }
    }

    public bool HasUnregisteredNoteActions => UnregisteredNoteActions.Count > 0;

    public string RegisterNoteActionsLabel => $"메모의 할일 {UnregisteredNoteActions.Count}건 등록";

    // ---- the event this todo is attached to ------------------------------------

    public bool HasLinkedEvent => _todo?.LinkedEventId is not null && LinkedEvent is not null;

    private CalendarEvent? LinkedEvent =>
        _todo?.LinkedEventId is { } id ? _repository.Events.FirstOrDefault(e => e.Id == id) : null;

    public string LinkedEventLabel =>
        LinkedEvent is { } linked
            ? $"→ {linked.Title} · {linked.Start.ToString(linked.IsAllDay ? "M'/'d (ddd)" : "M'/'d (ddd) HH:mm", Ko.Culture)}"
            : string.Empty;

    private void LoadAttached()
    {
        AttachedTodos.Clear();
        if (_event is null) return;

        var now = _clock();
        foreach (var todo in _repository.TodosAttachedTo(_event.Id))
        {
            AttachedTodos.Add(new AttachedTodoRow(todo, now, ToggleAttachedAsync, DetachAsync));
        }

        Raise(nameof(HasAttachedTodos));
        Raise(nameof(UnregisteredNoteActions));
        Raise(nameof(HasUnregisteredNoteActions));
        Raise(nameof(RegisterNoteActionsLabel));
    }

    private async Task AttachAsync()
    {
        if (_event is null || string.IsNullOrWhiteSpace(_newAttachedTitle)) return;

        await _repository.AttachTodoAsync(_event.Id, _newAttachedTitle, _clock());
        NewAttachedTitle = string.Empty;
        LoadAttached();
    }

    private async Task RegisterNoteActionsAsync()
    {
        if (_event is null) return;

        var pending = UnregisteredNoteActions;
        foreach (var title in pending) await _repository.AttachTodoAsync(_event.Id, title, _clock());

        Status = pending.Count == 0 ? string.Empty : $"할일 {pending.Count}건을 이 일정에 붙였습니다";
        LoadAttached();
    }

    private async Task ToggleAttachedAsync(AttachedTodoRow row)
    {
        await _repository.ToggleTodoAsync(row.Id, _clock());
        LoadAttached();
    }

    private async Task DetachAsync(AttachedTodoRow row)
    {
        await _repository.DetachTodoAsync(row.Id);
        LoadAttached();
    }

    private async Task DetachSelfAsync()
    {
        if (_todo is null) return;

        await _repository.DetachTodoAsync(_todo.Id);
        Raise(nameof(HasLinkedEvent));
        Raise(nameof(LinkedEventLabel));
    }

    // ---- the file this entry is about ---------------------------------------

    /// <summary>
    /// Whether to show the file row at all. An entry typed by hand names no file and gets
    /// the window exactly as it was before this existed, rather than a disabled button it
    /// can never use.
    /// </summary>
    public bool HasFile => _file is not null;

    public string FileName => _file is null ? string.Empty : SafeName(_file);

    /// <summary>The whole path, for the tooltip: the name alone does not say which copy.</summary>
    public string FilePath => _file ?? string.Empty;

    /// <summary>
    /// Opening is offered only for a file that is there and is not a program. Showing it in
    /// its folder stays available either way — that hands the decision to Explorer, where a
    /// person can see what they are about to run.
    /// </summary>
    public bool CanOpenFile => _file is not null && !AttachedFile.IsRunnable(_file) && Exists(_file);

    /// <summary>Says why the open button is off, and nothing at all when it is on.</summary>
    public string FileNote =>
        _file is null ? string.Empty
        : !Exists(_file) ? "파일을 찾을 수 없습니다. 옮겨졌거나 지워졌을 수 있습니다."
        : AttachedFile.IsRunnable(_file) ? "실행 파일은 Flowdeck에서 열지 않습니다. 폴더에서 확인해 주세요."
        : string.Empty;

    public bool HasFileNote => FileNote.Length > 0;

    public string Title
    {
        get => _title;
        set => Set(ref _title, value);
    }

    public string Notes
    {
        get => _notes;
        set
        {
            if (!Set(ref _notes, value)) return;

            Raise(nameof(UnregisteredNoteActions));
            Raise(nameof(HasUnregisteredNoteActions));
            Raise(nameof(RegisterNoteActionsLabel));
        }
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

        if (!saved.Found)
        {
            Status = "이미 삭제된 항목입니다";
            return;
        }

        // The edit is saved either way. What is left to say is whether the Outlook copy
        // came along, and that is only worth a word when it did not — so the window stays
        // open with the reason on it rather than closing on a change half made.
        var name = _repository.External?.DisplayName ?? "Outlook";
        switch (saved.External)
        {
            case ExternalSync.Lost:
                Status = $"{name} 쪽 항목을 찾지 못해 연결을 끊었습니다. 이 항목은 이제 로컬에만 있습니다.";
                return;

            case ExternalSync.Failed:
                Status = $"저장했지만 {name} 에는 반영하지 못했습니다.";
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
    /// Hands the file to whatever opens that kind of file. Checked again here rather than
    /// trusted from when the window opened: a file can go away while the window is up, and
    /// the failure to say so would be silent.
    /// </summary>
    private void OpenFile()
    {
        if (_file is null) return;

        if (!Exists(_file))
        {
            Status = "파일을 찾을 수 없습니다: " + _file;
            return;
        }

        if (AttachedFile.IsRunnable(_file))
        {
            Status = "실행 파일은 열지 않습니다.";
            return;
        }

        Launch(new ProcessStartInfo(_file) { UseShellExecute = true }, "파일을 열지 못했습니다");
    }

    /// <summary>
    /// Opens the folder with the file picked out in it. Falls back to the folder on its own
    /// when the file has gone, which is the more useful answer to "where did it go".
    /// </summary>
    private void ShowFile()
    {
        if (_file is null) return;

        if (Exists(_file))
        {
            Launch(
                new ProcessStartInfo("explorer.exe", $"/select,\"{_file}\"") { UseShellExecute = true },
                "폴더를 열지 못했습니다");
            return;
        }

        var folder = Folder(_file);
        if (folder is null || !Directory.Exists(folder))
        {
            Status = "폴더를 찾을 수 없습니다: " + _file;
            return;
        }

        Launch(new ProcessStartInfo(folder) { UseShellExecute = true }, "폴더를 열지 못했습니다");
    }

    private void Launch(ProcessStartInfo start, string failure)
    {
        try
        {
            Process.Start(start);
        }
        catch (Exception e) when (e is Win32Exception or InvalidOperationException or IOException)
        {
            Status = failure + ": " + e.Message;
        }
    }

    /// <summary>
    /// The path came out of a file another program wrote, so every one of these can be handed
    /// something that is not a path at all. None of them is worth an exception: a path that
    /// cannot be understood is simply a file that is not there.
    /// </summary>
    private static bool Exists(string path)
    {
        try
        {
            return File.Exists(path);
        }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    private static string SafeName(string path)
    {
        try
        {
            var name = Path.GetFileName(path);
            return name.Length > 0 ? name : path;
        }
        catch (ArgumentException)
        {
            return path;
        }
    }

    private static string? Folder(string path)
    {
        try
        {
            return Path.GetDirectoryName(path);
        }
        catch (Exception e) when (e is ArgumentException or PathTooLongException)
        {
            return null;
        }
    }

    /// <summary>
    /// Reads the two text boxes back into one instant. Refuses rather than guesses: a date
    /// that will not parse means saving would silently move the entry somewhere else.
    ///
    /// Emptying the box is not a failure to parse — it is the other way of saying "no date",
    /// alongside the checkbox. A todo takes that and becomes a someday item; an event still
    /// has to be told what day it is on.
    /// </summary>
    private bool TryBuildWhen(out DateTime? when)
    {
        when = null;
        if (!_hasDate) return true;

        if (string.IsNullOrWhiteSpace(_date)) return IsTodo;

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

/// <summary>One todo attached to the event being edited: tick it, or cut it loose.</summary>
public sealed class AttachedTodoRow
{
    public AttachedTodoRow(TodoItem todo, DateTime now, Func<AttachedTodoRow, Task> toggle, Func<AttachedTodoRow, Task> detach)
    {
        Id = todo.Id;
        Title = todo.Title;
        IsDone = todo.IsDone;
        IsOverdue = !todo.IsDone && todo.DueAt is { } due && due < now;
        DueLabel = todo.DueAt is { } at
            ? at.ToString(todo.HasTime ? "M'/'d (ddd) HH:mm" : "M'/'d (ddd)", Ko.Culture)
            : string.Empty;

        ToggleCommand = new AsyncRelayCommand(() => toggle(this));
        DetachCommand = new AsyncRelayCommand(() => detach(this));
    }

    public string Id { get; }

    public string Title { get; }

    public bool IsDone { get; }

    public bool IsOverdue { get; }

    public string DueLabel { get; }

    public ICommand ToggleCommand { get; }

    public ICommand DetachCommand { get; }
}

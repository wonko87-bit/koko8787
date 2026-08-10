using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Input;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Flowdeck.Mobile.Mvvm;

namespace Flowdeck.Mobile.ViewModels;

internal static class Ko
{
    public static readonly CultureInfo Culture = CultureInfo.GetCultureInfo("ko-KR");
}

/// <summary>One tag in the filter row, with how many entries carry it.</summary>
public sealed class TagChip : ObservableObject
{
    private bool _isSelected;

    public TagChip(string tag, int count, bool isSelected, Action<TagChip> onToggle)
    {
        Tag = tag;
        Label = "#" + tag;
        CountLabel = count.ToString(Ko.Culture);
        _isSelected = isSelected;
        ToggleCommand = new RelayCommand(() => onToggle(this));
    }

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

/// <summary>A day's worth of entries, or the tail that has no day.</summary>
public sealed class DayGroup : ObservableCollection<object>
{
    public DayGroup(string header, bool isToday, bool isPast)
    {
        Header = header;
        IsToday = isToday;
        IsPast = isPast;
    }

    public string Header { get; }

    public bool IsToday { get; }

    public bool IsPast { get; }
}

public sealed class TodoRow : ObservableObject
{
    private readonly TodoItem _item;

    public TodoRow(
        TodoItem item,
        DateTime now,
        Func<TodoRow, Task> toggle,
        Func<TodoRow, Task> delete,
        Action<TodoRow>? edit = null)
    {
        _item = item;
        IsOverdue = item.IsOverdue(now);

        ToggleCommand = new AsyncRelayCommand(() => toggle(this));
        DeleteCommand = new AsyncRelayCommand(() => delete(this));
        EditCommand = new RelayCommand(() => edit?.Invoke(this));
    }

    public string Id => _item.Id;

    public string Title => _item.Title;

    public bool IsDone => _item.IsDone;

    public bool IsOverdue { get; }

    public string DueLabel
    {
        get
        {
            if (_item.DueAt is null) return string.Empty;
            return _item.HasTime ? _item.DueAt.Value.ToString("tt h:mm", Ko.Culture) : "종일";
        }
    }

    public bool HasDue => _item.DueAt.HasValue;

    public string TagLabel => string.Join("  ", _item.Tags.Select(t => "#" + t));

    public bool HasTags => _item.Tags.Count > 0;

    public string PriorityLabel => _item.Priority switch
    {
        Priority.Urgent => "1",
        Priority.High => "2",
        Priority.Normal => "3",
        Priority.Low => "4",
        _ => string.Empty,
    };

    public bool HasPriority => _item.Priority != Priority.None;

    /// <summary>The first line of the note, so a row says whether there is more to read.</summary>
    public string NotePreview => FirstLine(_item.Notes);

    public bool HasNote => _item.Notes.Length > 0;

    public ICommand ToggleCommand { get; }

    public ICommand DeleteCommand { get; }

    public ICommand EditCommand { get; }

    internal static string FirstLine(string notes)
    {
        if (string.IsNullOrEmpty(notes)) return string.Empty;

        var end = notes.IndexOf('\n');
        var first = end < 0 ? notes : notes[..end];
        return end < 0 ? first : first + " …";
    }
}

public sealed class EventRow : ObservableObject
{
    public EventRow(EventOccurrence occurrence, Func<EventRow, Task> delete, Action<EventRow>? edit = null)
    {
        Id = occurrence.Source.Id;
        NotePreview = TodoRow.FirstLine(occurrence.Source.Notes);
        HasNote = occurrence.Source.Notes.Length > 0;
        EditCommand = new RelayCommand(() => edit?.Invoke(this));
        Title = occurrence.Title;
        TimeLabel = occurrence.IsAllDay ? "종일" : occurrence.Start.ToString("tt h:mm", Ko.Culture);
        TagLabel = string.Join("  ", occurrence.Source.Tags.Select(t => "#" + t));
        HasTags = occurrence.Source.Tags.Count > 0;

        DeleteCommand = new AsyncRelayCommand(() => delete(this));
    }

    public string Id { get; }

    public string Title { get; }

    public string TimeLabel { get; }

    public string TagLabel { get; }

    public bool HasTags { get; }

    public string NotePreview { get; }

    public bool HasNote { get; }

    public ICommand DeleteCommand { get; }

    public ICommand EditCommand { get; }
}

/// <summary>
/// A repeating entry, listed once as a rule rather than once per occurrence. Title first,
/// then the pattern in brackets — "3분할 운동 - 상체 [매주 월,목]".
/// </summary>
public sealed class RecurringRow : ObservableObject
{
    public RecurringRow(TodoItem todo, Func<RecurringRow, Task> delete)
    {
        Id = todo.Id;
        Title = todo.Title;
        PatternLabel = "[" + todo.Recurrence.DescribeShort(todo.DueAt) + "]";
        TagLabel = string.Join("  ", todo.Tags.Select(t => "#" + t));
        HasTags = todo.Tags.Count > 0;

        DeleteCommand = new AsyncRelayCommand(() => delete(this));
    }

    public RecurringRow(CalendarEvent source, Func<RecurringRow, Task> delete)
    {
        Id = source.Id;
        Title = source.Title;
        PatternLabel = "[" + source.Recurrence.DescribeShort(source.Start) + "]";
        TagLabel = string.Join("  ", source.Tags.Select(t => "#" + t));
        HasTags = source.Tags.Count > 0;

        DeleteCommand = new AsyncRelayCommand(() => delete(this));
    }

    public string Id { get; }

    public string Title { get; }

    public string PatternLabel { get; }

    public string TagLabel { get; }

    public bool HasTags { get; }

    public ICommand DeleteCommand { get; }
}

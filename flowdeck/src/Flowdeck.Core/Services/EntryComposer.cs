using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;

namespace Flowdeck.Core.Services;

/// <summary>What a single line of input produced.</summary>
public sealed class CaptureResult
{
    public TodoItem? Todo { get; init; }

    public CalendarEvent? Event { get; init; }

    public bool IsEmpty => Todo is null && Event is null;

    /// <summary>
    /// Set when the entry was saved locally but could not be copied to the external
    /// store. The capture itself still succeeded; only the copy did not.
    /// </summary>
    public string? ExternalError { get; set; }

    /// <summary>
    /// A link the shell should open once the entry is saved — today only the Teams
    /// scheduling form. Built here rather than launched here so nothing in Core has to
    /// know how a browser is started.
    /// </summary>
    public string? LaunchUrl { get; init; }
}

/// <summary>
/// Turns a <see cref="ParsedEntry"/> into the records it describes.
///
/// When the target is <see cref="EntryTarget.Both"/> the two records are cross-linked,
/// so ticking the todo off can also clear the matching calendar entry later.
/// </summary>
public static class EntryComposer
{
    public static CaptureResult Compose(ParsedEntry entry, DateTime now)
    {
        if (entry.IsEmpty) return new CaptureResult();

        var wantsTodo = entry.Target is EntryTarget.Todo or EntryTarget.Both;
        var wantsEvent = entry.Target is EntryTarget.Calendar or EntryTarget.Both;

        // An event has to sit somewhere on the calendar. With no date at all there is
        // nothing to draw, so the line stays a todo even if it asked for both.
        if (wantsEvent && !entry.HasDate)
        {
            wantsEvent = false;
            wantsTodo = true;
        }

        var todo = wantsTodo ? BuildTodo(entry, now) : null;
        var calendarEvent = wantsEvent ? BuildEvent(entry, now) : null;

        if (todo is not null && calendarEvent is not null)
        {
            todo.LinkedEventId = calendarEvent.Id;
            calendarEvent.LinkedTodoId = todo.Id;
        }

        return new CaptureResult
        {
            Todo = todo,
            Event = calendarEvent,
            LaunchUrl = TeamsDeepLink.NewMeetingFor(entry),
        };
    }

    private static TodoItem BuildTodo(ParsedEntry entry, DateTime now) => new()
    {
        Title = entry.Title,
        Notes = entry.Notes,
        DueAt = entry.Start,
        HasTime = entry.HasTime,
        Priority = entry.Priority,
        Tags = new List<string>(entry.Tags),
        Recurrence = entry.Recurrence.Clone(),
        ReminderMinutesBefore = entry.ReminderMinutesBefore,
        SourceText = entry.RawInput,
        CreatedAt = now,
        UpdatedAt = now,
    };

    private static CalendarEvent BuildEvent(ParsedEntry entry, DateTime now)
    {
        var start = entry.Start!.Value;
        var end = entry.End ?? (entry.HasTime ? start.AddHours(1) : start);

        // An all-day event covering a single day still needs a non-zero span,
        // otherwise nothing renders on the day strip.
        if (!entry.HasTime) end = end.Date.AddDays(1).AddSeconds(-1);

        return new CalendarEvent
        {
            Title = entry.Title,
            Notes = entry.Notes,
            Start = start,
            End = end,
            IsAllDay = !entry.HasTime,
            Tags = new List<string>(entry.Tags),
            Recurrence = entry.Recurrence.Clone(),
            ReminderMinutesBefore = entry.ReminderMinutesBefore,
            SourceText = entry.RawInput,
            CreatedAt = now,
            UpdatedAt = now,
        };
    }
}

namespace Flowdeck.Core.Models;

/// <summary>A single actionable item on the todo list.</summary>
public sealed class TodoItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Title { get; set; } = string.Empty;

    public string Notes { get; set; } = string.Empty;

    /// <summary>When the item is due. Null means "someday", no date attached.</summary>
    public DateTime? DueAt { get; set; }

    /// <summary>False when only a date was given, so the UI shows a day and not a clock time.</summary>
    public bool HasTime { get; set; }

    public Priority Priority { get; set; } = Priority.None;

    public List<string> Tags { get; set; } = new();

    public bool IsDone { get; set; }

    public DateTime? CompletedAt { get; set; }

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    public DateTime UpdatedAt { get; set; } = DateTime.Now;

    public Recurrence Recurrence { get; set; } = Recurrence.None;

    /// <summary>Set when the same input also produced a calendar event, so the two stay in step.</summary>
    public string? LinkedEventId { get; set; }

    /// <summary>The original text the user typed, kept for auditing and re-parsing.</summary>
    public string SourceText { get; set; } = string.Empty;

    /// <summary>Minutes before <see cref="DueAt"/> to raise a reminder. Null disables it.</summary>
    public int? ReminderMinutesBefore { get; set; }

    /// <summary>Set once the record has been copied to Outlook or another external store.</summary>
    public ExternalLink? ExternalLink { get; set; }

    public bool IsOverdue(DateTime now) => !IsDone && DueAt.HasValue && DueAt.Value < now;

    public TodoItem Clone() => (TodoItem)MemberwiseClone();
}

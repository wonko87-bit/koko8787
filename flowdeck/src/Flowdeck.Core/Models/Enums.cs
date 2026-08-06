namespace Flowdeck.Core.Models;

/// <summary>Where a parsed line of text should be filed.</summary>
public enum EntryTarget
{
    /// <summary>No explicit marker was given: file it in both places.</summary>
    Both = 0,

    /// <summary>Calendar only (<c>!CD</c>).</summary>
    Calendar = 1,

    /// <summary>Todo list only (<c>!TD</c>).</summary>
    Todo = 2,
}

/// <summary>Todo priority. Mirrors the familiar p1..p4 convention, p1 being the most urgent.</summary>
public enum Priority
{
    None = 0,
    Low = 1,
    Normal = 2,
    High = 3,
    Urgent = 4,
}

/// <summary>How an entry repeats.</summary>
public enum RecurrenceKind
{
    None = 0,
    Daily = 1,

    /// <summary>Monday through Friday.</summary>
    Weekdays = 2,
    Weekly = 3,
    Monthly = 4,
    Yearly = 5,
}

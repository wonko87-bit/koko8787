using Flowdeck.Core.Models;

namespace Flowdeck.Core.Parsing;

internal enum TokenKind
{
    Date,
    Time,
    Duration,
    Reminder,
    Recurrence,
    Priority,
    Tag,
    RangeSeparator,
    Noise,
}

/// <summary>
/// One recognised fragment of the input, together with the slice of text it came from.
/// The scanner produces a flat list of these; overlapping candidates are resolved once,
/// so every accepted token owns its characters outright and the leftover text is the title.
/// </summary>
internal sealed class Token
{
    public TokenKind Kind { get; init; }

    public int Start { get; set; }

    public int Length { get; set; }

    public int End => Start + Length;

    /// <summary>Tie-break when two candidates start at the same index and are the same length.</summary>
    public int Rank { get; init; }

    /// <summary>False for fragments that carry meaning but should stay in the title, e.g. "긴급".</summary>
    public bool Strip { get; init; } = true;

    public DateTime? Date { get; init; }

    public TimeFragment? Time { get; init; }

    public TimeSpan? Duration { get; init; }

    public int? ReminderMinutes { get; init; }

    public Recurrence? Recurrence { get; init; }

    public Priority Priority { get; init; }

    public string? Text { get; init; }
}

/// <summary>A clock time pulled out of the text, before it is attached to a date.</summary>
internal sealed class TimeFragment
{
    public int Hour { get; init; }

    public int Minute { get; init; }

    /// <summary>
    /// True when the text said am/pm outright ("오후 3시", "15:00"). A bare "3시" leaves this
    /// false, which lets a range like "3시~5시" infer that the end shares the start's half of the day.
    /// </summary>
    public bool ExplicitMeridiem { get; init; }
}

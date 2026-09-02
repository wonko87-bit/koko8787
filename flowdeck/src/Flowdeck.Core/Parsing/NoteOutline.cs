namespace Flowdeck.Core.Parsing;

/// <summary>What a line of a note is, by the word it opens with.</summary>
public enum NoteLineKind
{
    /// <summary>No keyword: the note's own words. The first of these is its summary.</summary>
    Text,

    // ---- shared between meetings and analysis runs ----

    /// <summary><c>결정:</c> — settled between people.</summary>
    Decision,

    /// <summary><c>할일:</c> — something somebody now has to do.</summary>
    Action,

    /// <summary><c>이슈:</c> — open, unresolved, or in the way.</summary>
    Issue,

    /// <summary><c>다음:</c> — the follow-up, the next meeting, the next step.</summary>
    Next,

    /// <summary><c>파일:</c> — a path. The same line the inbox hand-over writes.</summary>
    File,

    // ---- analysis runs ----

    /// <summary><c>조건:</c> — what the run was set up with.</summary>
    Condition,

    /// <summary><c>결과:</c> — the numbers that came out.</summary>
    Result,

    /// <summary><c>결론:</c> — what the numbers mean. Read off data, where a decision is made by people.</summary>
    Conclusion,
}

public sealed record NoteLine(NoteLineKind Kind, string Text);

/// <summary>
/// A note read by its line openers.
///
/// A note is prose, and stays prose: nothing here changes what is stored or how it is shown.
/// But a handful of words at the start of a line — <c>결정:</c>, <c>할일:</c>, <c>조건:</c> and
/// the rest — let the weekly report pull the decisions out of a week of meetings and the
/// conclusions out of a week of runs, without anyone filling in a form to get there. The
/// convention is deliberately small enough to remember and forgiving enough to mistype: a
/// space before the colon, a full-width colon, or a word that is not on the list all leave
/// the line as ordinary text.
/// </summary>
public sealed class NoteOutline
{
    private static readonly IReadOnlyDictionary<string, NoteLineKind> Openers =
        new Dictionary<string, NoteLineKind>(StringComparer.Ordinal)
        {
            ["결정"] = NoteLineKind.Decision,
            ["할일"] = NoteLineKind.Action,
            ["이슈"] = NoteLineKind.Issue,
            ["다음"] = NoteLineKind.Next,
            ["파일"] = NoteLineKind.File,
            ["조건"] = NoteLineKind.Condition,
            ["결과"] = NoteLineKind.Result,
            ["결론"] = NoteLineKind.Conclusion,
        };

    private static readonly char[] Colons = { ':', '：' };

    /// <summary>The longest opener, so a line cannot be an opener past this many characters in.</summary>
    private static readonly int LongestOpener = Openers.Keys.Max(k => k.Length);

    public static readonly NoteOutline Empty = new(Array.Empty<NoteLine>());

    private NoteOutline(IReadOnlyList<NoteLine> lines) => Lines = lines;

    /// <summary>Every non-blank line, in the order written.</summary>
    public IReadOnlyList<NoteLine> Lines { get; }

    /// <summary>The first line without a keyword, which is what the note is about in a phrase.</summary>
    public string Summary => Text.FirstOrDefault() ?? string.Empty;

    public IReadOnlyList<string> Text => Of(NoteLineKind.Text);

    public IReadOnlyList<string> Decisions => Of(NoteLineKind.Decision);

    public IReadOnlyList<string> Actions => Of(NoteLineKind.Action);

    public IReadOnlyList<string> Issues => Of(NoteLineKind.Issue);

    public IReadOnlyList<string> Next => Of(NoteLineKind.Next);

    public IReadOnlyList<string> Files => Of(NoteLineKind.File);

    public IReadOnlyList<string> Conditions => Of(NoteLineKind.Condition);

    public IReadOnlyList<string> Results => Of(NoteLineKind.Result);

    public IReadOnlyList<string> Conclusions => Of(NoteLineKind.Conclusion);

    /// <summary>Whether any line carries a keyword at all — a note that is only prose has no outline.</summary>
    public bool HasStructure => Lines.Any(l => l.Kind != NoteLineKind.Text);

    /// <summary>Whether the note describes a run: at least one of 조건/결과/결론 is present.</summary>
    public bool IsAnalysis => Lines.Any(l => l.Kind is NoteLineKind.Condition or NoteLineKind.Result or NoteLineKind.Conclusion);

    public static NoteOutline Parse(string? notes)
    {
        if (string.IsNullOrWhiteSpace(notes)) return Empty;

        var lines = new List<NoteLine>();

        foreach (var raw in notes.Split('\n'))
        {
            var line = raw.Trim();
            if (line.Length == 0) continue;

            lines.Add(Classify(line));
        }

        return lines.Count == 0 ? Empty : new NoteOutline(lines);
    }

    /// <summary>
    /// A line is an opener line when everything before its first colon, trimmed, is one of
    /// the words — so "결정 :" counts and "결정 사항: 없음" does not, and a colon further along
    /// the line ("회의: 3시부터" as prose) is not looked at, because the word in front of it
    /// would not be on the list anyway.
    /// </summary>
    private static NoteLine Classify(string line)
    {
        var colon = line.IndexOfAny(Colons);

        // Bounded so a long line is not scanned for a colon that could not be an opener's.
        if (colon <= 0 || colon > LongestOpener + 2) return new NoteLine(NoteLineKind.Text, line);

        var head = line[..colon].Trim();
        if (!Openers.TryGetValue(head, out var kind)) return new NoteLine(NoteLineKind.Text, line);

        var body = line[(colon + 1)..].Trim();

        // "결정:" and nothing after it says nothing. Treated as the prose it looks like rather
        // than dropped, so a person who meant to come back to it still sees it.
        return body.Length == 0 ? new NoteLine(NoteLineKind.Text, line) : new NoteLine(kind, body);
    }

    private IReadOnlyList<string> Of(NoteLineKind kind) =>
        Lines.Where(l => l.Kind == kind).Select(l => l.Text).ToList();
}

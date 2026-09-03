using System.Globalization;
using System.Text;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;

namespace Flowdeck.Core.Services;

/// <summary>An entry named in the report with, where it has one, the day it belongs to.</summary>
public sealed record ReportItem(string Title, DateTime? When, bool HasTime, string? File);

/// <summary>An entry whose note is read out in the report: a meeting, or an analysis run.</summary>
public sealed record ReportNote(string Title, DateTime When, bool HasTime, NoteOutline Outline);

/// <summary>One line pulled out of a note, with the entry it came from.</summary>
public sealed record ReportLine(string Source, string Text);

/// <summary>
/// A week of the workspace, read back as the material for a weekly report.
///
/// Nothing new is recorded to make this: everything here was already written down in the
/// course of the week — todos ticked off, notes left on meetings, the lines a run's note
/// opened with. The report gathers it into the shape a status report wants, and hands over
/// text to be edited and pasted, not a finished document. The sentences are the writer's.
///
/// Built and rendered apart, so what goes into each section can be checked without reading
/// prose back.
/// </summary>
public sealed class WeeklyReport
{
    private static readonly CultureInfo Korean = CultureInfo.GetCultureInfo("ko-KR");

    /// <summary>How far past the week to look for what is coming.</summary>
    private const int LookAheadDays = 7;

    private WeeklyReport(DateTime from, DateTime to)
    {
        From = from;
        To = to;
    }

    /// <summary>First day, at midnight.</summary>
    public DateTime From { get; }

    /// <summary>Last day, at midnight — inclusive.</summary>
    public DateTime To { get; }

    /// <summary>Todos ticked off during the week, in the order they were.</summary>
    public IReadOnlyList<ReportItem> Completed { get; private init; } = Array.Empty<ReportItem>();

    /// <summary>
    /// Meetings with something written on them: copies taken from Outlook, and any other
    /// event whose note uses the line openers. A meeting nobody wrote anything about is not
    /// reported — there is nothing to report.
    /// </summary>
    public IReadOnlyList<ReportNote> Meetings { get; private init; } = Array.Empty<ReportNote>();

    /// <summary>Entries whose note describes a run: at least one of 조건/결과/결론.</summary>
    public IReadOnlyList<ReportNote> Runs { get; private init; } = Array.Empty<ReportNote>();

    /// <summary>
    /// Every 결정: line across the week — from meetings, from runs, and from any plain todo
    /// whose note used the openers. Where a line was written does not change what it is.
    /// </summary>
    public IReadOnlyList<ReportLine> Decisions { get; private init; } = Array.Empty<ReportLine>();

    /// <summary>Every 할일: line across the week, from the same places.</summary>
    public IReadOnlyList<ReportLine> Actions { get; private init; } = Array.Empty<ReportLine>();

    public IReadOnlyList<ReportLine> Issues { get; private init; } = Array.Empty<ReportLine>();

    public IReadOnlyList<ReportLine> NextSteps { get; private init; } = Array.Empty<ReportLine>();

    /// <summary>Files named by the week's entries — what was handled, by name.</summary>
    public IReadOnlyList<string> Files { get; private init; } = Array.Empty<string>();

    /// <summary>Open todos that were due by the end of the week and are not done.</summary>
    public IReadOnlyList<ReportItem> Unfinished { get; private init; } = Array.Empty<ReportItem>();

    /// <summary>Open todos due in the week after.</summary>
    public IReadOnlyList<ReportItem> Upcoming { get; private init; } = Array.Empty<ReportItem>();

    public bool IsEmpty =>
        Completed.Count + Meetings.Count + Runs.Count + Files.Count + Unfinished.Count + Upcoming.Count == 0;

    /// <summary>The week a day falls in, given which day weeks start on.</summary>
    public static (DateTime From, DateTime To) WeekOf(DateTime day, DayOfWeek firstDay)
    {
        var offset = ((int)day.DayOfWeek - (int)firstDay + 7) % 7;
        var from = day.Date.AddDays(-offset);
        return (from, from.AddDays(6));
    }

    public static WeeklyReport Build(WorkspaceRepository repository, DateTime from, DateTime to)
    {
        from = from.Date;
        to = to.Date;
        if (to < from) (from, to) = (to, from);

        var end = to.AddDays(1);
        bool InWeek(DateTime? at) => at.HasValue && at.Value >= from && at.Value < end;

        // ---- todos ------------------------------------------------------------

        var completed = repository.Todos
            .Where(t => t.IsDone && InWeek(t.CompletedAt))
            .OrderBy(t => t.CompletedAt)
            .Select(t => new ReportItem(t.Title, t.CompletedAt, false, AttachedFile.PathIn(t.Notes)))
            .ToList();

        var runs = new List<ReportNote>();
        var meetings = new List<ReportNote>();
        var files = new List<string>();

        // Plain todos whose note uses the openers. Not meetings and not runs, so they get
        // no block of their own — but a decision written on a phone call is still one of
        // the week's decisions, and goes into the lists across the week with the rest.
        var asides = new List<ReportNote>();

        foreach (var todo in repository.Todos)
        {
            // A todo belongs to the week it was finished in, failing that the week it was
            // due, failing that the week it was last touched.
            var anchor = todo.CompletedAt ?? todo.DueAt ?? todo.UpdatedAt;
            var thisWeek = InWeek(anchor);

            var outline = NoteOutline.Parse(todo.Notes);
            if (thisWeek && outline.IsAnalysis)
            {
                runs.Add(new ReportNote(todo.Title, anchor, todo.HasTime && todo.DueAt == anchor, outline));
            }
            else if (thisWeek && outline.HasStructure)
            {
                asides.Add(new ReportNote(todo.Title, anchor, false, outline));
            }

            if (InWeek(todo.CreatedAt) || InWeek(todo.CompletedAt) || thisWeek)
            {
                files.AddRange(outline.Files);
            }
        }

        // ---- events -----------------------------------------------------------

        foreach (var occurrence in repository.OccurrencesBetween(from, to))
        {
            var source = occurrence.Source;
            var outline = NoteOutline.Parse(source.Notes);
            var note = new ReportNote(occurrence.Title, occurrence.Start, !occurrence.IsAllDay, outline);

            // A copy taken from Outlook is a meeting whatever its note says; anything else
            // is a meeting if it was written up like one. A run is what is left.
            if (source.Origin is not null || (outline.HasStructure && !outline.IsAnalysis))
            {
                if (outline.Lines.Count > 0) meetings.Add(note);
            }
            else if (outline.IsAnalysis)
            {
                runs.Add(note);
            }

            files.AddRange(outline.Files);
        }

        meetings.Sort((a, b) => a.When.CompareTo(b.When));
        runs.Sort((a, b) => a.When.CompareTo(b.When));

        // ---- across the week ------------------------------------------------------

        // Meetings first, then runs, then the rest, each in time order: the lists read in
        // the order the week's notes were taken, not in the order the entries were stored.
        asides.Sort((a, b) => a.When.CompareTo(b.When));
        var noted = meetings.Concat(runs).Concat(asides).ToList();

        var decisions = noted.SelectMany(n => n.Outline.Decisions.Select(d => new ReportLine(n.Title, d))).ToList();
        var actions = noted.SelectMany(n => n.Outline.Actions.Select(a => new ReportLine(n.Title, a))).ToList();
        var issues = noted.SelectMany(n => n.Outline.Issues.Select(i => new ReportLine(n.Title, i))).ToList();
        var next = noted.SelectMany(n => n.Outline.Next.Select(x => new ReportLine(n.Title, x))).ToList();

        // ---- what is open ---------------------------------------------------------

        var open = repository.OpenTodos();

        var unfinished = open
            .Where(t => t.DueAt.HasValue && t.DueAt.Value < end)
            .OrderBy(t => t.DueAt)
            .Select(t => new ReportItem(t.Title, t.DueAt, t.HasTime, AttachedFile.PathIn(t.Notes)))
            .ToList();

        var lookAhead = end.AddDays(LookAheadDays);
        var upcoming = open
            .Where(t => t.DueAt.HasValue && t.DueAt.Value >= end && t.DueAt.Value < lookAhead)
            .OrderBy(t => t.DueAt)
            .Select(t => new ReportItem(t.Title, t.DueAt, t.HasTime, AttachedFile.PathIn(t.Notes)))
            .ToList();

        return new WeeklyReport(from, to)
        {
            Completed = completed,
            Meetings = meetings,
            Runs = runs,
            Decisions = decisions,
            Actions = actions,
            Issues = issues,
            NextSteps = next,
            Files = files.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            Unfinished = unfinished,
            Upcoming = upcoming,
        };
    }

    // ---- text -----------------------------------------------------------------

    public string Render()
    {
        var text = new StringBuilder();

        text.Append("주간보고 초안 · ")
            .Append(From.ToString("yyyy-MM-dd (ddd)", Korean))
            .Append(" ~ ")
            .AppendLine(To.ToString("yyyy-MM-dd (ddd)", Korean));

        if (IsEmpty)
        {
            text.AppendLine().AppendLine("이 기간에 기록된 항목이 없습니다.");
            return text.ToString();
        }

        Items("완료", Completed);
        Notes("회의", Meetings);
        Lines("이번 주 결정", Decisions);
        Lines("액션 아이템", Actions);
        Lines("이슈", Issues);
        Lines("다음", NextSteps);
        Notes("해석", Runs);

        if (Files.Count > 0)
        {
            Heading("다룬 자료", Files.Count);
            foreach (var file in Files) text.Append("  - ").AppendLine(FileName(file));
        }

        Items("미완료", Unfinished);
        Items("다음 주", Upcoming);

        return text.ToString().TrimEnd() + Environment.NewLine;

        void Heading(string title, int count) =>
            text.AppendLine().Append("■ ").Append(title).Append(" (").Append(count).AppendLine(")");

        void Items(string title, IReadOnlyList<ReportItem> items)
        {
            if (items.Count == 0) return;

            Heading(title, items.Count);
            foreach (var item in items)
            {
                text.Append("  - ").Append(item.Title);
                if (item.When.HasValue) text.Append(" · ").Append(Day(item.When.Value, item.HasTime));
                if (item.File is not null) text.Append(" · ").Append(FileName(item.File));
                text.AppendLine();
            }
        }

        void Notes(string title, IReadOnlyList<ReportNote> notes)
        {
            if (notes.Count == 0) return;

            Heading(title, notes.Count);
            foreach (var note in notes)
            {
                text.Append("  ").Append(note.Title).Append(" · ").AppendLine(Day(note.When, note.HasTime));
                foreach (var line in note.Outline.Lines)
                {
                    text.Append("    ").AppendLine(Restore(line));
                }
            }
        }

        void Lines(string title, IReadOnlyList<ReportLine> lines)
        {
            if (lines.Count == 0) return;

            Heading(title, lines.Count);
            foreach (var line in lines)
            {
                text.Append("  - [").Append(line.Source).Append("] ").AppendLine(line.Text);
            }
        }
    }

    // The slash is quoted: bare, it is the culture's date separator, which in Korean is a
    // dot and a space, and "8. 20" is not what anyone here writes.
    private static string Day(DateTime when, bool hasTime) =>
        when.ToString(hasTime ? "M'/'d (ddd) HH:mm" : "M'/'d (ddd)", Korean);

    /// <summary>
    /// The last segment of a path, whichever way its slashes lean. Not Path.GetFileName: the
    /// path was written on a Windows machine, and this library also runs on a phone, where
    /// a backslash is an ordinary character and the whole path would come back as the name.
    /// </summary>
    private static string FileName(string path)
    {
        var trimmed = path.TrimEnd('\\', '/');
        var cut = trimmed.LastIndexOfAny(new[] { '\\', '/' });
        var name = cut < 0 ? trimmed : trimmed[(cut + 1)..];
        return name.Length > 0 ? name : path;
    }

    /// <summary>A note line as it was written, opener and all, so the block reads as the note did.</summary>
    private static string Restore(NoteLine line) => line.Kind switch
    {
        NoteLineKind.Decision => "결정: " + line.Text,
        NoteLineKind.Action => "할일: " + line.Text,
        NoteLineKind.Issue => "이슈: " + line.Text,
        NoteLineKind.Next => "다음: " + line.Text,
        NoteLineKind.File => "파일: " + FileName(line.Text),
        NoteLineKind.Condition => "조건: " + line.Text,
        NoteLineKind.Result => "결과: " + line.Text,
        NoteLineKind.Conclusion => "결론: " + line.Text,
        _ => line.Text,
    };
}

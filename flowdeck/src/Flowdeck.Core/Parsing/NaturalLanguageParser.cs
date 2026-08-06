using System.Text;
using System.Text.RegularExpressions;
using Flowdeck.Core.Models;

namespace Flowdeck.Core.Parsing;

/// <summary>
/// Reads a free-text line and works out what it means: where it belongs, when it happens,
/// how often it repeats, and what is left over as the title.
/// </summary>
public sealed class NaturalLanguageParser
{
    /// <summary>Tags that double as a routing marker, e.g. "#할일".</summary>
    private static readonly Dictionary<string, EntryTarget> TagTargets =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["할일"] = EntryTarget.Todo,
            ["투두"] = EntryTarget.Todo,
            ["todo"] = EntryTarget.Todo,
            ["td"] = EntryTarget.Todo,
            ["일정"] = EntryTarget.Calendar,
            ["캘린더"] = EntryTarget.Calendar,
            ["cal"] = EntryTarget.Calendar,
            ["cd"] = EntryTarget.Calendar,
        };

    private static readonly Regex Whitespace = new(@"\s+", RegexOptions.CultureInvariant);
    private static readonly char[] TrimChars = { ' ', '\t', ',', '.', '-', '–', '—', ':', ';', '·', '~' };

    private readonly RoutingRules _rules;

    public NaturalLanguageParser(RoutingRules? rules = null) => _rules = rules ?? new RoutingRules();

    public RoutingRules Rules => _rules;

    /// <summary>Default event length when a start time is given but no end.</summary>
    public TimeSpan DefaultEventDuration { get; set; } = TimeSpan.FromHours(1);

    /// <summary>
    /// When only a clock time is given, roll it to tomorrow if that time has already
    /// passed today. "9시 운동" typed late at night means tomorrow morning.
    /// </summary>
    public bool RollBareTimeToTomorrow { get; set; } = true;

    /// <summary>
    /// Reads a bare "2시" as 14:00 rather than 02:00, which is how times are quoted in
    /// Korean scheduling. Only hours 1–7 are shifted: "9시" stays in the morning, and
    /// anything written with 오전/오후 is left exactly as typed. Turn off for literal hours.
    /// </summary>
    public bool AssumeAfternoonForBareHours { get; set; } = true;

    /// <summary>Bare hours at or below this are read as afternoon when the rule is on.</summary>
    private const int AfternoonCutoff = 7;

    public ParsedEntry Parse(string input, DateTime now)
    {
        var entry = new ParsedEntry { RawInput = input ?? string.Empty };
        if (string.IsNullOrWhiteSpace(input)) return entry;

        var explicitTarget = _rules.FindMarker(input, out var body);
        var tokens = TemporalScanner.Scan(body, now);

        ApplyTokens(entry, tokens, now, ref explicitTarget);

        entry.Title = BuildTitle(body, tokens);
        if (entry.Title.Length == 0) entry.Title = ParsedEntry.UntitledPlaceholder;

        entry.TargetWasExplicit = explicitTarget.HasValue;
        entry.Target = explicitTarget ?? _rules.ClassifyByKeyword(body);

        return entry;
    }

    private void ApplyTokens(
        ParsedEntry entry,
        List<Token> tokens,
        DateTime now,
        ref EntryTarget? explicitTarget)
    {
        var dates = new List<DateTime>();
        var times = new List<TimeFragment>();
        TimeSpan? duration = null;

        foreach (var token in tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.Date when token.Date.HasValue:
                    dates.Add(token.Date.Value);
                    break;

                case TokenKind.Time when token.Time is not null:
                    times.Add(token.Time);
                    break;

                case TokenKind.Duration when token.Duration.HasValue:
                    duration ??= token.Duration.Value;
                    break;

                case TokenKind.Reminder:
                    entry.ReminderMinutesBefore ??= token.ReminderMinutes;
                    break;

                case TokenKind.Recurrence when token.Recurrence is not null:
                    if (!entry.Recurrence.IsRepeating) entry.Recurrence = token.Recurrence.Clone();
                    break;

                case TokenKind.Priority:
                    if (token.Priority > entry.Priority) entry.Priority = token.Priority;
                    break;

                case TokenKind.Tag when token.Text is not null:
                    if (!explicitTarget.HasValue && TagTargets.TryGetValue(token.Text, out var tagTarget))
                    {
                        explicitTarget = tagTarget;
                    }
                    else if (!entry.Tags.Contains(token.Text, StringComparer.OrdinalIgnoreCase))
                    {
                        entry.Tags.Add(token.Text);
                    }

                    break;
            }
        }

        ResolveSchedule(entry, dates, times, duration, now);

        // "매주 화요일" carries its weekday in the date, not in the repeat word.
        if (entry.Recurrence.Kind == RecurrenceKind.Weekly
            && entry.Recurrence.DaysOfWeek.Count == 0
            && entry.Start.HasValue)
        {
            entry.Recurrence.DaysOfWeek.Add(entry.Start.Value.DayOfWeek);
        }
    }

    private void ResolveSchedule(
        ParsedEntry entry,
        List<DateTime> dates,
        List<TimeFragment> times,
        TimeSpan? duration,
        DateTime now)
    {
        if (dates.Count == 0 && times.Count == 0) return;

        var startDate = dates.Count > 0 ? dates[0] : now.Date;
        var endDate = dates.Count > 1 ? dates[1] : startDate;

        if (times.Count == 0)
        {
            entry.HasTime = false;
            entry.Start = startDate;
            entry.End = endDate;
            return;
        }

        entry.HasTime = true;
        var startHour = EffectiveHour(times[0]);
        var start = startDate.AddHours(startHour).AddMinutes(times[0].Minute);

        if (dates.Count == 0 && RollBareTimeToTomorrow && start < now)
        {
            start = start.AddDays(1);
            endDate = endDate.AddDays(1);
        }

        entry.Start = start;

        if (times.Count > 1)
        {
            var endFragment = times[1];
            var endHour = EffectiveHour(endFragment);
            var end = endDate.AddHours(endHour).AddMinutes(endFragment.Minute);

            // "10시~2시": a bare end time shares the start's half of the day.
            if (!endFragment.ExplicitMeridiem && end < start && endHour < 12)
            {
                end = end.AddHours(12);
            }

            if (end < start) end = end.AddDays(1);
            entry.End = end;
        }
        else
        {
            entry.End = start + (duration ?? DefaultEventDuration);
        }
    }

    private int EffectiveHour(TimeFragment fragment)
    {
        if (fragment.ExplicitMeridiem || !AssumeAfternoonForBareHours) return fragment.Hour;
        return fragment.Hour is >= 1 and <= AfternoonCutoff ? fragment.Hour + 12 : fragment.Hour;
    }

    /// <summary>Rebuilds the title from whatever text the scanner did not claim.</summary>
    private static string BuildTitle(string body, List<Token> tokens)
    {
        var builder = new StringBuilder(body.Length);
        var cursor = 0;

        foreach (var token in tokens.Where(t => t.Strip).OrderBy(t => t.Start))
        {
            if (token.Start > cursor) builder.Append(body, cursor, token.Start - cursor);
            builder.Append(' ');
            cursor = Math.Max(cursor, token.End);
        }

        if (cursor < body.Length) builder.Append(body, cursor, body.Length - cursor);

        var title = Whitespace.Replace(builder.ToString(), " ").Trim();
        return title.Trim(TrimChars).Trim();
    }
}

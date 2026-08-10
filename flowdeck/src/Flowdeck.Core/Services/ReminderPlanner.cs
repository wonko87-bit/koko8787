using System.Globalization;

namespace Flowdeck.Core.Services;

/// <summary>One reminder that wants to go off, and what it should say when it does.</summary>
public sealed record DueReminder(string Key, string Title, DateTime At, DateTime Subject, bool IsTodo, bool HasTime)
{
    /// <summary>
    /// Stable across recalculations so a platform scheduler can replace an alarm rather
    /// than stack a second one on top of it. A repeating event needs the occurrence in the
    /// key: the same entry comes round again and again, and each round is its own alarm.
    /// </summary>
    public int AlarmId => Key.GetHashCode() & 0x7FFFFFFF;
}

/// <summary>
/// Works out which reminders are coming and when — the part of notifying that has nothing
/// to do with any particular platform, and therefore the part worth testing.
///
/// A phone schedules alarms from this; the desktop polls it. Neither has to agree with the
/// other about what "due" means, because neither decides.
/// </summary>
public static class ReminderPlanner
{
    /// <summary>
    /// Reminders falling between now and <paramref name="horizon"/> later, soonest first.
    ///
    /// Bounded because a phone has a limited number of alarms worth holding and a repeating
    /// entry would otherwise produce them without end. Whatever is dropped comes back on
    /// the next pass, by which time it is nearer.
    /// </summary>
    public static IReadOnlyList<DueReminder> Upcoming(
        WorkspaceRepository repository,
        DateTime from,
        TimeSpan horizon,
        int max = 20)
    {
        if (repository is null || max <= 0) return Array.Empty<DueReminder>();

        var until = from + horizon;
        var found = new List<DueReminder>();

        foreach (var todo in repository.OpenTodos())
        {
            if (todo.ReminderMinutesBefore is not > 0 || todo.DueAt is null) continue;

            var at = todo.DueAt.Value.AddMinutes(-todo.ReminderMinutesBefore.Value);
            if (at < from || at > until) continue;

            found.Add(new DueReminder(
                $"todo:{todo.Id}:{todo.DueAt.Value:O}",
                todo.Title,
                at,
                todo.DueAt.Value,
                IsTodo: true,
                todo.HasTime));
        }

        // Expanded a day either side of the window: an occurrence just outside it can still
        // carry a reminder that falls inside.
        foreach (var occurrence in repository.OccurrencesBetween(from.Date.AddDays(-1), until.Date.AddDays(1)))
        {
            var minutes = occurrence.Source.ReminderMinutesBefore;
            if (minutes is not > 0) continue;

            var at = occurrence.Start.AddMinutes(-minutes.Value);
            if (at < from || at > until) continue;

            found.Add(new DueReminder(
                $"event:{occurrence.Source.Id}:{occurrence.Start:O}",
                occurrence.Title,
                at,
                occurrence.Start,
                IsTodo: false,
                !occurrence.IsAllDay));
        }

        return found.OrderBy(r => r.At).Take(max).ToList();
    }

    /// <summary>
    /// Reminders that came due in the window just passed, for a scheduler waking up to ask
    /// what it missed. <paramref name="grace"/> bounds how late is still worth saying: a
    /// reminder surfacing hours after the fact is noise, not a reminder.
    /// </summary>
    public static IReadOnlyList<DueReminder> Overdue(
        WorkspaceRepository repository,
        DateTime now,
        TimeSpan grace)
    {
        if (repository is null) return Array.Empty<DueReminder>();

        return Upcoming(repository, now - grace, grace, max: int.MaxValue)
            .Where(r => r.At <= now)
            .ToList();
    }

    private static readonly CultureInfo Korean = CultureInfo.GetCultureInfo("ko-KR");

    /// <summary>The line under the title in a notification.</summary>
    public static string Describe(DueReminder reminder) =>
        reminder.HasTime
            ? reminder.Subject.ToString("M월 d일 tt h:mm", Korean)
            : reminder.Subject.ToString("M월 d일", Korean);
}

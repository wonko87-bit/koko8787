using System.Globalization;
using System.Windows.Threading;
using Flowdeck.Core.Services;

namespace Flowdeck.Windows.Services;

/// <summary>
/// Watches for entries whose reminder has come due and raises a tray notification.
///
/// Reminders fire at most once per run of the app. Nothing is written back to the
/// workspace, so a reminder that was missed while the machine was asleep is simply
/// skipped rather than firing late in a burst.
/// </summary>
public sealed class ReminderService : IDisposable
{
    private static readonly CultureInfo Korean = CultureInfo.GetCultureInfo("ko-KR");

    /// <summary>How stale a due reminder may be before it is dropped instead of shown.</summary>
    private static readonly TimeSpan Grace = TimeSpan.FromMinutes(10);

    private readonly WorkspaceRepository _repository;
    private readonly Action<string, string> _notify;
    private readonly Func<DateTime> _clock;
    private readonly HashSet<string> _fired = new();
    private readonly DispatcherTimer _timer;

    public ReminderService(
        WorkspaceRepository repository,
        Action<string, string> notify,
        Func<DateTime>? clock = null)
    {
        _repository = repository;
        _notify = notify;
        _clock = clock ?? (() => DateTime.Now);

        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _timer.Tick += (_, _) => Poll();
    }

    public void Start() => _timer.Start();

    /// <summary>Exposed so the check can be driven directly in a test.</summary>
    public void Poll()
    {
        var now = _clock();

        foreach (var todo in _repository.Todos)
        {
            if (todo.IsDone || todo.ReminderMinutesBefore is null || todo.DueAt is null) continue;

            var at = todo.DueAt.Value.AddMinutes(-todo.ReminderMinutesBefore.Value);
            if (ShouldFire("todo:" + todo.Id, at, now))
            {
                _notify("할일 알림", $"{todo.Title} · {Describe(todo.DueAt.Value, todo.HasTime)}");
            }
        }

        // Only the next few days are expanded: a reminder further out cannot be due yet.
        foreach (var occurrence in _repository.OccurrencesBetween(now.Date, now.Date.AddDays(2)))
        {
            var minutes = occurrence.Source.ReminderMinutesBefore;
            if (minutes is null) continue;

            var at = occurrence.Start.AddMinutes(-minutes.Value);
            if (ShouldFire($"event:{occurrence.Source.Id}:{occurrence.Start:O}", at, now))
            {
                _notify("일정 알림", $"{occurrence.Title} · {Describe(occurrence.Start, !occurrence.IsAllDay)}");
            }
        }
    }

    private bool ShouldFire(string key, DateTime at, DateTime now)
    {
        if (at > now || now - at > Grace) return false;

        return _fired.Add(key);
    }

    private static string Describe(DateTime when, bool hasTime) =>
        hasTime ? when.ToString("M월 d일 tt h:mm", Korean) : when.ToString("M월 d일", Korean);

    public void Dispose() => _timer.Stop();
}

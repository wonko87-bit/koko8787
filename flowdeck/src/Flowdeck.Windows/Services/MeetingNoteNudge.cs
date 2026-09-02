using System.Windows.Threading;
using Flowdeck.Core.Integration;
using Flowdeck.Core.Services;
using Flowdeck.Core.Settings;

namespace Flowdeck.Windows.Services;

/// <summary>
/// Asks for a line on a meeting the moment it ends.
///
/// Notes on meetings do not get written later; they get written now or not at all. So this
/// watches for a meeting reaching its end and puts up one balloon — click it and the meeting
/// is taken in (if it was not already) and opened on its note. Ignore it and it goes away.
///
/// Only meetings from the outside calendar are watched, taken in or not: those are the ones
/// with no note field of their own. A copy that already carries a note is not asked about
/// again. Like the reminder service, this fires at most once per meeting per run, and a
/// meeting that ended while the machine was asleep is let go rather than announced late.
/// </summary>
public sealed class MeetingNoteNudge : IDisposable
{
    private static readonly TimeSpan Grace = TimeSpan.FromMinutes(10);

    private readonly WorkspaceRepository _repository;
    private readonly ExternalCalendarFeed _calendar;
    private readonly Func<MeetingNudgeScope> _scope;
    private readonly Action<string, string, Action> _notify;
    private readonly Action<ExternalOccurrence> _takeInAndOpen;
    private readonly Action<string> _open;
    private readonly Func<DateTime> _clock;
    private readonly HashSet<string> _fired = new();
    private readonly DispatcherTimer _timer;

    public MeetingNoteNudge(
        WorkspaceRepository repository,
        ExternalCalendarFeed calendar,
        Func<MeetingNudgeScope> scope,
        Action<string, string, Action> notify,
        Action<ExternalOccurrence> takeInAndOpen,
        Action<string> open,
        Func<DateTime>? clock = null)
    {
        _repository = repository;
        _calendar = calendar;
        _scope = scope;
        _notify = notify;
        _takeInAndOpen = takeInAndOpen;
        _open = open;
        _clock = clock ?? (() => DateTime.Now);

        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _timer.Tick += (_, _) => Poll();
    }

    public void Start() => _timer.Start();

    public void Poll()
    {
        var scope = _scope();
        if (scope == MeetingNudgeScope.Off) return;

        var now = _clock();

        // Copies already taken. Asked about whether or not the overlay is on, since they
        // are the user's own entries now.
        foreach (var copy in _repository.Events)
        {
            if (copy.Origin is null || copy.IsAllDay) continue;
            if (copy.Notes.Trim().Length > 0) continue;

            if (Due("copy:" + copy.Id, copy.End, now))
            {
                var id = copy.Id;
                _notify("회의 메모", Ask(copy.Title), () => _open(id));
            }
        }

        if (scope != MeetingNudgeScope.All) return;

        // Everything else on the overlay. Yesterday is looked at too, for the meeting that
        // ran past midnight; the grace window keeps anything older from firing.
        foreach (var day in new[] { now.Date.AddDays(-1), now.Date })
        {
            foreach (var occurrence in _calendar.On(day, _repository.Hides))
            {
                // A leave day or a public holiday ends at midnight and is not a meeting.
                if (occurrence.IsAllDay) continue;

                if (Due($"meeting:{occurrence.EntryId}:{occurrence.Start:O}", occurrence.End, now))
                {
                    var target = occurrence;
                    _notify("회의 메모", Ask(occurrence.Title), () => _takeInAndOpen(target));
                }
            }
        }
    }

    private static string Ask(string title) =>
        $"방금 끝난 '{(string.IsNullOrWhiteSpace(title) ? "(제목 없음)" : title)}', 한 줄 남길까요? 누르면 메모창이 열립니다.";

    private bool Due(string key, DateTime at, DateTime now)
    {
        if (at > now || now - at > Grace) return false;

        return _fired.Add(key);
    }

    public void Dispose() => _timer.Stop();
}

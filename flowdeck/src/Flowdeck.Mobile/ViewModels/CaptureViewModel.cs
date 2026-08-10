using System.Windows.Input;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Mobile.Mvvm;
using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile.ViewModels;

/// <summary>
/// The capture screen: one box, and a running account of how the line is being read.
///
/// The whole reading is <c>Flowdeck.Core</c>'s, unchanged from the desktop — same dates,
/// same markers, same tags. What is different here is only that the preview has to earn its
/// space on a small screen, so each part of it appears only once there is something to say.
/// </summary>
public sealed class CaptureViewModel : ObservableObject
{
    private readonly Func<DateTime> _clock;

    private string _input = string.Empty;
    private ParsedEntry _preview = new();
    private EntryTarget? _forcedTarget;
    private string _status = string.Empty;

    public CaptureViewModel(Func<DateTime>? clock = null)
    {
        _clock = clock ?? (() => DateTime.Now);

        SubmitCommand = new AsyncRelayCommand(SubmitAsync, () => HasInput);
        ForceTodoCommand = new RelayCommand(() => ForceTarget(EntryTarget.Todo));
        ForceCalendarCommand = new RelayCommand(() => ForceTarget(EntryTarget.Calendar));
        ForceBothCommand = new RelayCommand(() => ForceTarget(EntryTarget.Both));
    }

    public string Input
    {
        get => _input;
        set
        {
            if (!Set(ref _input, value)) return;

            _preview = Workspace.Parser.Parse(_input, _clock());
            Status = string.Empty;
            RaisePreview();
        }
    }

    public bool HasInput => !string.IsNullOrWhiteSpace(_input);

    public ICommand SubmitCommand { get; }

    public ICommand ForceTodoCommand { get; }

    public ICommand ForceCalendarCommand { get; }

    public ICommand ForceBothCommand { get; }

    public string TargetLabel => Effective().DescribeTarget();

    public string TitlePreview => HasInput ? Effective().Title : string.Empty;

    public string SchedulePreview => HasInput ? Effective().DescribeSchedule() : string.Empty;

    public string RecurrenceLabel => Effective().Recurrence.Describe();

    public bool HasRecurrence => Effective().Recurrence.IsRepeating;

    public string TagLabel => string.Join("  ", Effective().Tags.Select(t => "#" + t));

    public bool HasTags => Effective().Tags.Count > 0;

    public string PriorityLabel => Effective().Priority switch
    {
        Priority.Urgent => "우선순위 1",
        Priority.High => "우선순위 2",
        Priority.Normal => "우선순위 3",
        Priority.Low => "우선순위 4",
        _ => string.Empty,
    };

    public bool HasPriority => Effective().Priority != Priority.None;

    public string ReminderLabel
    {
        get
        {
            var minutes = Effective().ReminderMinutesBefore;
            if (minutes is null) return string.Empty;
            return minutes.Value % 60 == 0 && minutes.Value >= 60
                ? $"{minutes.Value / 60}시간 전 알림"
                : $"{minutes.Value}분 전 알림";
        }
    }

    public bool HasReminder => Effective().ReminderMinutesBefore.HasValue;

    /// <summary>True when the line asked for a Teams meeting, which opens the Teams app.</summary>
    public bool OpensTeamsMeeting => Effective().OpenTeamsMeeting;

    public bool IsTodoForced => Effective().Target == EntryTarget.Todo;

    public bool IsCalendarForced => Effective().Target == EntryTarget.Calendar;

    public bool IsBothForced => Effective().Target == EntryTarget.Both;

    /// <summary>What just happened, shown under the button and cleared on the next keystroke.</summary>
    public string Status
    {
        get => _status;
        set
        {
            if (Set(ref _status, value)) Raise(nameof(HasStatus));
        }
    }

    public bool HasStatus => _status.Length > 0;

    /// <summary>
    /// Overrides the routing for this one entry. Choosing what is already chosen releases
    /// the override, which is the only way back to the standing default without retyping.
    /// </summary>
    public void ForceTarget(EntryTarget target)
    {
        _forcedTarget = _forcedTarget == target ? null : target;
        RaisePreview();
    }

    public async Task SubmitAsync()
    {
        if (!HasInput) return;

        var entry = Effective();
        var result = await Workspace.Repository.CaptureAsync(entry, _clock());
        if (result.IsEmpty) return;

        var saved = entry.Title;
        Input = string.Empty;
        _forcedTarget = null;
        RaisePreview();
        Status = $"저장했습니다 · {saved}";

        // The meeting form is a link, and a link opens the Teams app just as well from a
        // phone. Failing to open it must not lose the entry, which is already saved.
        if (result.LaunchUrl is not null)
        {
            try
            {
                await Launcher.OpenAsync(result.LaunchUrl);
            }
            catch (Exception)
            {
                Status = "저장했습니다. Teams를 열지 못했습니다";
            }
        }
    }

    private ParsedEntry Effective()
    {
        if (_forcedTarget is null || _preview.IsEmpty) return _preview;

        return new ParsedEntry
        {
            RawInput = _preview.RawInput,
            Title = _preview.Title,
            Target = _forcedTarget.Value,
            TargetWasExplicit = true,
            Start = _preview.Start,
            End = _preview.End,
            HasTime = _preview.HasTime,
            Priority = _preview.Priority,
            Tags = _preview.Tags,
            Recurrence = _preview.Recurrence,
            ReminderMinutesBefore = _preview.ReminderMinutesBefore,
            PushExternal = _preview.PushExternal,
            OpenTeamsMeeting = _preview.OpenTeamsMeeting,
            Attendees = _preview.Attendees,
        };
    }

    private void RaisePreview()
    {
        Raise(nameof(Input));
        Raise(nameof(HasInput));
        Raise(nameof(TargetLabel));
        Raise(nameof(TitlePreview));
        Raise(nameof(SchedulePreview));
        Raise(nameof(RecurrenceLabel));
        Raise(nameof(HasRecurrence));
        Raise(nameof(TagLabel));
        Raise(nameof(HasTags));
        Raise(nameof(PriorityLabel));
        Raise(nameof(HasPriority));
        Raise(nameof(ReminderLabel));
        Raise(nameof(HasReminder));
        Raise(nameof(OpensTeamsMeeting));
        Raise(nameof(IsTodoForced));
        Raise(nameof(IsCalendarForced));
        Raise(nameof(IsBothForced));
        (SubmitCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
    }
}

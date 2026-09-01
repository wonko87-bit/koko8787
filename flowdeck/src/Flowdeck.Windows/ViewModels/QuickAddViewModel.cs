using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>
/// Backs the hot-key capture window. Everything it shows is derived from the raw
/// text, so the user can see exactly how the line was read before committing it.
/// </summary>
public sealed class QuickAddViewModel : ObservableObject
{
    private readonly WorkspaceRepository _repository;
    private readonly NaturalLanguageParser _parser;
    private readonly Func<DateTime> _clock;

    private string _input = string.Empty;
    private ParsedEntry _preview = new();
    private EntryTarget? _forcedTarget;
    private bool? _forcedExternal;
    private bool? _forcedTeams;

    public QuickAddViewModel(
        WorkspaceRepository repository,
        NaturalLanguageParser parser,
        Func<DateTime>? clock = null)
    {
        _repository = repository;
        _parser = parser;
        _clock = clock ?? (() => DateTime.Now);
    }

    /// <summary>Raised once an entry has been filed, so the window can close itself.</summary>
    public event EventHandler<CaptureResult>? Captured;

    public string Input
    {
        get => _input;
        set
        {
            if (!Set(ref _input, value)) return;
            Reparse();
        }
    }

    public bool HasInput => !string.IsNullOrWhiteSpace(_input);

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

    /// <summary>
    /// How many lines of note are being written. The note itself is not repeated here — it
    /// is already on screen, in the box being typed into. What the preview has to confirm is
    /// that it was read as a note, and so is not going to end up in the title.
    /// </summary>
    public string NoteLabel
    {
        get
        {
            var note = Effective().Notes;
            if (note.Length == 0) return string.Empty;

            var lines = note.Split('\n').Length;
            return lines > 1 ? $"메모 {lines}줄" : "메모";
        }
    }

    public bool HasNote => Effective().Notes.Length > 0;

    /// <summary>
    /// True when this entry will also be written into Outlook. Reports what will actually
    /// happen, not what was asked for: with no Outlook on the machine, or the integration
    /// switched off, <c>!OL</c> is going nowhere and the preview should not claim otherwise.
    /// </summary>
    public bool PushesToOutlook => Effective().PushExternal && _repository.External is not null;

    /// <summary>True when the target is settled rather than guessed from the words.</summary>
    public bool TargetIsPinned =>
        _forcedTarget.HasValue
        || _preview.TargetWasExplicit
        || _parser.Rules.DefaultTarget != TargetDefault.Auto;

    /// <summary>
    /// Overrides the routing for this one entry, for when the standing default is not what
    /// was wanted. Choosing the target that is already forced releases the override and
    /// hands the decision back to the parser.
    /// </summary>
    public void ForceTarget(EntryTarget target)
    {
        _forcedTarget = _forcedTarget == target ? null : target;
        RaisePreviewProperties();
    }

    /// <summary>Flips the Outlook flag for this one entry: the keyboard twin of !OL / !NOL.</summary>
    public void ToggleExternal()
    {
        _forcedExternal = !(_forcedExternal ?? _preview.PushExternal);
        RaisePreviewProperties();
    }

    /// <summary>
    /// Flips the Teams flag for this one entry: the keyboard twin of !TM. Re-reads the
    /// line rather than patching the result, because turning a meeting on is what turns an
    /// <c>@이름</c> from a word in the title into someone to invite.
    /// </summary>
    public void ToggleTeams()
    {
        _forcedTeams = !(_forcedTeams ?? _preview.OpenTeamsMeeting);
        Reparse();
    }

    /// <summary>True when saving will also open the Teams scheduling form.</summary>
    public bool OpensTeamsMeeting => Effective().OpenTeamsMeeting;

    /// <summary>Who the meeting will be addressed to, for the preview line.</summary>
    public string AttendeeLabel => string.Join(", ", Effective().Attendees);

    public bool HasAttendees => Effective().Attendees.Count > 0;

    public void Reset() => Reset(null);

    /// <summary>
    /// Clears the overlay, optionally starting it off with text taken from somewhere else.
    ///
    /// Text picked up from another window is a starting point rather than an entry: a date or
    /// a !TD is usually still wanted, and line breaks inside a selection mean nothing to the
    /// parser, so they become the spaces they stood in for.
    /// </summary>
    public void Reset(string? seed)
    {
        _forcedTarget = null;
        _forcedExternal = null;
        _forcedTeams = null;

        Input = string.IsNullOrWhiteSpace(seed)
            ? string.Empty
            : string.Join(" ", seed.Split('\n', '\r', StringSplitOptions.RemoveEmptyEntries))
                .Replace("  ", " ", StringComparison.Ordinal)
                .Trim();
    }

    public async Task<bool> SubmitAsync()
    {
        if (!HasInput) return false;

        var entry = Effective();
        var result = await _repository.CaptureAsync(entry, _clock());
        if (result.IsEmpty) return false;

        Captured?.Invoke(this, result);
        Reset();
        return true;
    }

    private void Reparse()
    {
        _preview = _parser.Parse(_input, _clock(), _forcedTeams);
        RaisePreviewProperties();
    }

    /// <summary>The parse with any manual override folded in.</summary>
    private ParsedEntry Effective()
    {
        // The Teams flag is not folded in here: it went through the parse, so the preview
        // already carries it along with the attendees it pulled out.
        if (_preview.IsEmpty) return _preview;
        if (_forcedTarget is null && _forcedExternal is null) return _preview;

        return new ParsedEntry
        {
            RawInput = _preview.RawInput,
            Title = _preview.Title,
            Target = _forcedTarget ?? _preview.Target,
            TargetWasExplicit = _forcedTarget.HasValue || _preview.TargetWasExplicit,
            Start = _preview.Start,
            End = _preview.End,
            HasTime = _preview.HasTime,
            Priority = _preview.Priority,
            Tags = _preview.Tags,
            Recurrence = _preview.Recurrence,
            ReminderMinutesBefore = _preview.ReminderMinutesBefore,
            PushExternal = _forcedExternal ?? _preview.PushExternal,
            OpenTeamsMeeting = _preview.OpenTeamsMeeting,
            Attendees = _preview.Attendees,
        };
    }

    private void RaisePreviewProperties()
    {
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
        Raise(nameof(NoteLabel));
        Raise(nameof(HasNote));
        Raise(nameof(PushesToOutlook));
        Raise(nameof(OpensTeamsMeeting));
        Raise(nameof(AttendeeLabel));
        Raise(nameof(HasAttendees));
        Raise(nameof(TargetIsPinned));
    }
}

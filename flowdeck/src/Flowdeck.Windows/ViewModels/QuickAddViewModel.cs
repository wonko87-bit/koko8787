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

    /// <summary>True when this entry will also be written into Outlook.</summary>
    public bool PushesToOutlook => Effective().PushExternal;

    /// <summary>True when the target came from a marker rather than from a keyword guess.</summary>
    public bool TargetIsPinned => _forcedTarget.HasValue || _preview.TargetWasExplicit;

    /// <summary>
    /// Overrides the routing for this one entry, for when the keyword sweep guessed wrong.
    /// Passing null hands the decision back to the parser.
    /// </summary>
    public void ForceTarget(EntryTarget? target)
    {
        _forcedTarget = target;
        RaisePreviewProperties();
    }

    public void Reset()
    {
        _forcedTarget = null;
        Input = string.Empty;
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
        _preview = _parser.Parse(_input, _clock());
        RaisePreviewProperties();
    }

    /// <summary>The parse with any manual target override folded in.</summary>
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
        Raise(nameof(PushesToOutlook));
        Raise(nameof(TargetIsPinned));
    }
}

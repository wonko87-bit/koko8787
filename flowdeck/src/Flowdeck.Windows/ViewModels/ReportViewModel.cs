using System.Globalization;
using System.Windows.Input;
using Flowdeck.Core.Services;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>
/// The weekly report window's state: which week, and the text made from it.
///
/// The text is a draft in a box, not a document. It is rebuilt whenever the week changes and
/// otherwise left alone, so what the person has typed into it survives until they ask for a
/// fresh one. Copying and saving are the window's business — they need the clipboard and a
/// file dialog, which do not belong here.
/// </summary>
public sealed class ReportViewModel : ObservableObject
{
    private readonly WorkspaceRepository _repository;
    private readonly DayOfWeek _firstDay;
    private readonly Func<DateTime> _clock;

    private DateTime _from;
    private DateTime _to;
    private bool _isThisWeek = true;
    private bool _isLastWeek;
    private bool _isCustom;
    private string _fromText = string.Empty;
    private string _toText = string.Empty;
    private string _text = string.Empty;
    private string _status = string.Empty;

    public ReportViewModel(WorkspaceRepository repository, DayOfWeek firstDay, Func<DateTime>? clock = null)
    {
        _repository = repository;
        _firstDay = firstDay;
        _clock = clock ?? (() => DateTime.Now);

        RebuildCommand = new RelayCommand(Rebuild);

        ChooseThisWeek();
    }

    public ICommand RebuildCommand { get; }

    public bool IsThisWeek
    {
        get => _isThisWeek;
        set
        {
            if (Set(ref _isThisWeek, value) && value) ChooseThisWeek();
        }
    }

    public bool IsLastWeek
    {
        get => _isLastWeek;
        set
        {
            if (Set(ref _isLastWeek, value) && value) ChooseLastWeek();
        }
    }

    /// <summary>
    /// Any two dates. The boxes start off showing the week that was selected, so "custom"
    /// is a small edit rather than typing two dates from nothing.
    /// </summary>
    public bool IsCustom
    {
        get => _isCustom;
        set
        {
            if (Set(ref _isCustom, value) && value)
            {
                // Set through the fields, not the properties: the boxes are being filled in
                // with the week already shown, and that is not a reason to build it again.
                _fromText = _from.ToString("yyyy-MM-dd", Ko.Culture);
                _toText = _to.ToString("yyyy-MM-dd", Ko.Culture);
                Raise(nameof(FromText));
                Raise(nameof(ToText));
                Status = "날짜를 고치면 바로 다시 만들어집니다";
            }
        }
    }

    /// <summary>
    /// The boxes rebuild as they are typed in, the moment both hold a date. Typing a date is
    /// already the request; a button to make it take is a step to forget.
    /// </summary>
    public string FromText
    {
        get => _fromText;
        set
        {
            if (Set(ref _fromText, value)) RebuildIfDatesTyped();
        }
    }

    public string ToText
    {
        get => _toText;
        set
        {
            if (Set(ref _toText, value)) RebuildIfDatesTyped();
        }
    }

    private void RebuildIfDatesTyped()
    {
        if (!_isCustom) return;

        if (DateTime.TryParse(_fromText, Ko.Culture, DateTimeStyles.None, out _)
            && DateTime.TryParse(_toText, Ko.Culture, DateTimeStyles.None, out _))
        {
            Rebuild();
        }
    }

    public string RangeLabel =>
        $"{_from.ToString("yyyy-MM-dd (ddd)", Ko.Culture)} ~ {_to.ToString("yyyy-MM-dd (ddd)", Ko.Culture)}";

    /// <summary>What the file is called when saved: the week, so a folder of them sorts itself.</summary>
    public string SuggestedFileName => $"주간보고_{_from:yyyyMMdd}-{_to:yyyyMMdd}.txt";

    public string Text
    {
        get => _text;
        set => Set(ref _text, value);
    }

    public string Status
    {
        get => _status;
        set
        {
            if (Set(ref _status, value)) Raise(nameof(HasStatus));
        }
    }

    public bool HasStatus => _status.Length > 0;

    private void ChooseThisWeek()
    {
        (_from, _to) = WeeklyReport.WeekOf(_clock(), _firstDay);
        Rebuild();
    }

    private void ChooseLastWeek()
    {
        (_from, _to) = WeeklyReport.WeekOf(_clock().AddDays(-7), _firstDay);
        Rebuild();
    }

    /// <summary>
    /// Makes the text again for the chosen week. In custom mode the boxes are read first,
    /// and refused rather than guessed at when they do not hold a date.
    /// </summary>
    public void Rebuild()
    {
        if (_isCustom)
        {
            if (!DateTime.TryParse(_fromText, Ko.Culture, DateTimeStyles.None, out var from)
                || !DateTime.TryParse(_toText, Ko.Culture, DateTimeStyles.None, out var to))
            {
                Status = "날짜는 2026-08-17 처럼 써 주세요";
                return;
            }

            _from = from.Date;
            _to = to.Date;
            if (_to < _from) (_from, _to) = (_to, _from);
        }

        var report = WeeklyReport.Build(_repository, _from, _to);
        Text = report.Render();
        Status = report.IsEmpty ? "이 기간에는 기록된 항목이 없습니다" : string.Empty;

        Raise(nameof(RangeLabel));
        Raise(nameof(SuggestedFileName));
    }
}

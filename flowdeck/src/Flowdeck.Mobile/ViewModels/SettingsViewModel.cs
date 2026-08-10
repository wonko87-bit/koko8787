using System.Windows.Input;
using Flowdeck.Core.Export;
using Flowdeck.Core.Parsing;
using Flowdeck.Mobile.Mvvm;
using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile.ViewModels;

public sealed class SettingsViewModel : ObservableObject
{
    private string _status = string.Empty;

    public SettingsViewModel()
    {
        ImportCommand = new AsyncRelayCommand(ImportAsync);
        DefaultAutoCommand = new RelayCommand(() => DefaultTarget = TargetDefault.Auto);
        DefaultTodoCommand = new RelayCommand(() => DefaultTarget = TargetDefault.Todo);
        DefaultCalendarCommand = new RelayCommand(() => DefaultTarget = TargetDefault.Calendar);
        DefaultBothCommand = new RelayCommand(() => DefaultTarget = TargetDefault.Both);
    }

    public ICommand ImportCommand { get; }

    public ICommand DefaultAutoCommand { get; }

    public ICommand DefaultTodoCommand { get; }

    public ICommand DefaultCalendarCommand { get; }

    public ICommand DefaultBothCommand { get; }

    /// <summary>
    /// Where an unmarked line goes. Named anything but Auto, the keyword sweep is skipped —
    /// the same rule as the desktop, because a phone that guessed differently would be worse
    /// than one that guessed at all.
    /// </summary>
    public TargetDefault DefaultTarget
    {
        get
        {
            var rules = Workspace.Settings.Routing;
            if (rules.DefaultTarget != TargetDefault.Auto) return rules.DefaultTarget;
            return rules.UseKeywordHints ? TargetDefault.Auto : TargetDefault.Both;
        }

        set
        {
            Workspace.Settings.Routing.DefaultTarget = value;
            Workspace.Settings.Routing.UseKeywordHints = value == TargetDefault.Auto;
            Persist();
            Raise();
            Raise(nameof(IsAuto));
            Raise(nameof(IsTodo));
            Raise(nameof(IsCalendar));
            Raise(nameof(IsBoth));
            Raise(nameof(DefaultTargetHint));
        }
    }

    public bool IsAuto => DefaultTarget == TargetDefault.Auto;

    public bool IsTodo => DefaultTarget == TargetDefault.Todo;

    public bool IsCalendar => DefaultTarget == TargetDefault.Calendar;

    public bool IsBoth => DefaultTarget == TargetDefault.Both;

    public string DefaultTargetHint => DefaultTarget switch
    {
        TargetDefault.Todo => "!TD 를 매번 붙이지 않아도 됩니다",
        TargetDefault.Calendar => "!CD 를 매번 붙이지 않아도 됩니다",
        TargetDefault.Both => "표기가 없으면 항상 양쪽에 등록합니다",
        _ => "'회의'는 일정, '제출'은 할일로. 판단이 서지 않으면 양쪽 모두",
    };

    public bool AfternoonBias
    {
        get => Workspace.Settings.AssumeAfternoonForBareHours;
        set
        {
            if (Workspace.Settings.AssumeAfternoonForBareHours == value) return;

            Workspace.Settings.AssumeAfternoonForBareHours = value;
            Persist();
            Raise();
        }
    }

    public bool WeekStartsMonday
    {
        get => Workspace.Settings.FirstDayOfWeek != DayOfWeek.Sunday;
        set
        {
            var wanted = value ? DayOfWeek.Monday : DayOfWeek.Sunday;
            if (Workspace.Settings.FirstDayOfWeek == wanted) return;

            Workspace.Settings.FirstDayOfWeek = wanted;
            Persist();
            Raise();
        }
    }

    public string DataFolder => Workspace.DataFolder;

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
    /// Reads a file exported by another copy of Flowdeck. Additive: nothing already here is
    /// replaced or removed, and the same file twice adds nothing the second time.
    /// </summary>
    public async Task ImportAsync()
    {
        try
        {
            var picked = await FilePicker.Default.PickAsync(new PickOptions { PickerTitle = "Flowdeck 파일 고르기" });
            if (picked is null) return;

            await using var stream = await picked.OpenReadAsync();
            using var reader = new StreamReader(stream);
            var archive = TransferFile.Read(await reader.ReadToEndAsync());

            var result = await Workspace.Repository.ImportAsync(archive);
            Status = result.Describe();
        }
        catch (FormatException e)
        {
            Status = e.Message;
        }
        catch (Exception e)
        {
            // The picker can fail in a dozen platform-specific ways — a permission refused,
            // a provider that died, a URI that no longer resolves. None of them is worth
            // taking the app down for.
            Status = "파일을 열지 못했습니다: " + e.Message;
        }
    }

    private static void Persist()
    {
        Workspace.ApplyParserSettings();
        Workspace.SaveSettings();
    }
}

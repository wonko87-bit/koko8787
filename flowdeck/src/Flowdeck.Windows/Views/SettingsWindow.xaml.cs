using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Flowdeck.Core.Export;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Settings;
using Flowdeck.Windows.Interop;
using Flowdeck.Windows.ViewModels;

namespace Flowdeck.Windows.Views;

public partial class SettingsWindow : Window
{
    private readonly IAppShell _shell;

    /// <summary>Set while the controls are being populated, so handlers do not fire mid-load.</summary>
    private bool _loading;

    public SettingsWindow(IAppShell shell)
    {
        InitializeComponent();

        _shell = shell;
        Load();
    }

    private AppSettings Settings => _shell.Settings;

    private void Load()
    {
        _loading = true;

        ThemeDark.IsChecked = Settings.Theme == AppTheme.Dark;
        ThemeLight.IsChecked = Settings.Theme == AppTheme.Light;

        PinDesktop.IsChecked = Settings.PinMode == WidgetPinMode.Desktop;
        PinTop.IsChecked = Settings.PinMode == WidgetPinMode.AlwaysOnTop;
        PinNormal.IsChecked = Settings.PinMode == WidgetPinMode.Normal;

        Opacity60.IsChecked = Settings.WidgetOpacity < 0.7;
        Opacity80.IsChecked = Settings.WidgetOpacity is >= 0.7 and < 0.87;
        Opacity92.IsChecked = Settings.WidgetOpacity is >= 0.87 and < 0.97;
        Opacity100.IsChecked = Settings.WidgetOpacity >= 0.97;

        WeekStartMonday.IsChecked = Settings.FirstDayOfWeek != DayOfWeek.Sunday;
        WeekStartSunday.IsChecked = Settings.FirstDayOfWeek == DayOfWeek.Sunday;
        ColorWeekends.IsChecked = Settings.ColorWeekends;

        QuickAddHotkey.Text = Settings.QuickAddHotkey;
        ToggleWidgetHotkey.Text = Settings.ToggleWidgetHotkey;
        AgendaEventsHotkey.Text = Settings.AgendaEventsHotkey;
        AgendaTodosHotkey.Text = Settings.AgendaTodosHotkey;
        CaptureSelectionHotkey.Text = Settings.CaptureSelectionHotkey;

        LoadDefaultTarget();

        // The section stays hidden where Outlook is absent: a switch that could never do
        // anything is worse than no switch at all.
        OutlookSection.Visibility = _shell.Repository.ExternalDetected
            ? Visibility.Visible
            : Visibility.Collapsed;
        UseOutlook.IsChecked = Settings.EnableOutlook;
        PushToOutlook.IsChecked = Settings.PushToOutlookByDefault;
        PushToOutlook.IsEnabled = Settings.EnableOutlook;
        ShowOutlookCalendar.IsChecked = Settings.ShowOutlookCalendar;
        ShowOutlookCalendar.IsEnabled = Settings.EnableOutlook;
        LoadRefreshInterval();

        AfternoonBias.IsChecked = Settings.AssumeAfternoonForBareHours;

        DescribeContacts();

        LaunchAtStartup.IsChecked = StartupService.Current() != StartupState.Off;
        DescribeStartup();
        ShowWidgetOnStart.IsChecked = Settings.ShowWidgetOnStart;

        WatchInbox.IsChecked = Settings.EnableInboxWatch;
        DescribeInbox();

        _loading = false;

        AfternoonBias.Checked += OnParserOptionChanged;
        AfternoonBias.Unchecked += OnParserOptionChanged;
        UseOutlook.Checked += OnOutlookChanged;
        UseOutlook.Unchecked += OnOutlookChanged;
        PushToOutlook.Checked += OnOutlookChanged;
        PushToOutlook.Unchecked += OnOutlookChanged;
        ColorWeekends.Checked += OnWeekStartChanged;
        ColorWeekends.Unchecked += OnWeekStartChanged;
        LaunchAtStartup.Checked += OnStartupChanged;
        LaunchAtStartup.Unchecked += OnStartupChanged;
        ShowWidgetOnStart.Checked += OnShowOnStartChanged;
        ShowWidgetOnStart.Unchecked += OnShowOnStartChanged;
        WatchInbox.Checked += OnWatchInboxChanged;
        WatchInbox.Unchecked += OnWatchInboxChanged;
    }

    private void OnThemeChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.Theme = ThemeLight.IsChecked == true ? AppTheme.Light : AppTheme.Dark;
        _shell.ApplyTheme(Settings.Theme);
        _shell.SaveSettings();
    }

    private void OnPinChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.PinMode = PinTop.IsChecked == true
            ? WidgetPinMode.AlwaysOnTop
            : PinNormal.IsChecked == true
                ? WidgetPinMode.Normal
                : WidgetPinMode.Desktop;

        _shell.ApplyWidgetSettings();
        _shell.SaveSettings();
    }

    private void OnOpacityChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.WidgetOpacity = Opacity60.IsChecked == true ? 0.60
            : Opacity80.IsChecked == true ? 0.80
            : Opacity92.IsChecked == true ? 0.92
            : 1.0;

        _shell.ApplyWidgetSettings();
        _shell.SaveSettings();
    }

    /// <summary>
    /// The week-start choice feeds the parser as well as the grid, so that "이번주 일요일"
    /// lands on the Sunday the user can actually see in the calendar.
    /// </summary>
    private void OnWeekStartChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.FirstDayOfWeek = WeekStartSunday.IsChecked == true ? DayOfWeek.Sunday : DayOfWeek.Monday;
        Settings.ColorWeekends = ColorWeekends.IsChecked == true;

        _shell.ApplyParserSettings();
        _shell.ApplyWidgetSettings();
        _shell.SaveSettings();
    }

    /// <summary>
    /// Shows the stored choice. A settings file written before this option existed carries
    /// only the keyword switch, and "off" meant "always both" — which is now one of the
    /// four choices, so it maps straight across with nothing to migrate.
    /// </summary>
    private void LoadDefaultTarget()
    {
        var rules = Settings.Routing;
        var choice = rules.DefaultTarget != TargetDefault.Auto ? rules.DefaultTarget
            : rules.UseKeywordHints ? TargetDefault.Auto
            : TargetDefault.Both;

        DefaultAuto.IsChecked = choice == TargetDefault.Auto;
        DefaultTodo.IsChecked = choice == TargetDefault.Todo;
        DefaultCalendar.IsChecked = choice == TargetDefault.Calendar;
        DefaultBoth.IsChecked = choice == TargetDefault.Both;
        DescribeDefaultTarget(choice);
    }

    private void OnDefaultTargetChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        var choice = DefaultTodo.IsChecked == true ? TargetDefault.Todo
            : DefaultCalendar.IsChecked == true ? TargetDefault.Calendar
            : DefaultBoth.IsChecked == true ? TargetDefault.Both
            : TargetDefault.Auto;

        Settings.Routing.DefaultTarget = choice;

        // Kept in step so the two never contradict each other in the file: the keyword
        // sweep is exactly the Auto choice, under its old name.
        Settings.Routing.UseKeywordHints = choice == TargetDefault.Auto;

        DescribeDefaultTarget(choice);
        _shell.ApplyParserSettings();
        _shell.SaveSettings();
    }

    private void DescribeDefaultTarget(TargetDefault choice) =>
        DefaultTargetHint.Text = choice switch
        {
            TargetDefault.Todo => "!TD 를 매번 붙이지 않아도 됩니다. 예외는 !CD 또는 Ctrl+2",
            TargetDefault.Calendar => "!CD 를 매번 붙이지 않아도 됩니다. 예외는 !TD 또는 Ctrl+1",
            TargetDefault.Both => "표기가 없으면 항상 일정과 할일 양쪽에 등록합니다",
            _ => "'회의'는 일정, '제출'은 할일로. 판단이 서지 않으면 양쪽 모두",
        };

    private void OnOutlookChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.EnableOutlook = UseOutlook.IsChecked == true;
        Settings.PushToOutlookByDefault = PushToOutlook.IsChecked == true;

        _shell.Repository.ExternalEnabled = Settings.EnableOutlook;
        PushToOutlook.IsEnabled = Settings.EnableOutlook;
        ShowOutlookCalendar.IsEnabled = Settings.EnableOutlook;

        _shell.ApplyParserSettings();
        _shell.ApplyCalendarOverlay();
        _shell.SaveSettings();
    }

    // ---- the calendar overlay ----------------------------------------------

    /// <summary>
    /// Lights the preset matching the stored interval, or falls back to the custom box when
    /// it is a number nobody offered — which is exactly what the box is for.
    /// </summary>
    private void LoadRefreshInterval()
    {
        var minutes = Settings.EffectiveOutlookRefreshMinutes;

        Refresh5.IsChecked = minutes == 5;
        Refresh10.IsChecked = minutes == 10;
        Refresh30.IsChecked = minutes == 30;
        Refresh60.IsChecked = minutes == 60;
        RefreshCustom.IsChecked = minutes is not (5 or 10 or 30 or 60);

        RefreshMinutes.Text = minutes.ToString(Ko.Culture);
        DescribeRefresh();
    }

    private void DescribeRefresh()
    {
        var minutes = Settings.EffectiveOutlookRefreshMinutes;

        RefreshRow.IsEnabled = ShowOutlookCalendar.IsChecked == true && Settings.EnableOutlook;
        RefreshHint.Text = minutes < 5
            ? $"{minutes}분마다. 짧으면 Outlook을 그만큼 자주 깨웁니다"
            : $"{minutes}분마다 다시 읽습니다";
    }

    private void OnCalendarOverlayChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.ShowOutlookCalendar = ShowOutlookCalendar.IsChecked == true;
        DescribeRefresh();

        _shell.ApplyCalendarOverlay();
        _shell.SaveSettings();
    }

    private void OnRefreshPresetChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        // The custom button chooses nothing by itself; it hands the decision to the box.
        if (RefreshCustom.IsChecked == true)
        {
            RefreshMinutes.Focus();
            RefreshMinutes.SelectAll();
            return;
        }

        Settings.OutlookCalendarRefreshMinutes =
            Refresh5.IsChecked == true ? 5
            : Refresh30.IsChecked == true ? 30
            : Refresh60.IsChecked == true ? 60
            : 10;

        RefreshMinutes.Text = Settings.OutlookCalendarRefreshMinutes.ToString(Ko.Culture);
        DescribeRefresh();

        _shell.ApplyCalendarOverlay();
        _shell.SaveSettings();
    }

    /// <summary>
    /// Takes whatever was typed, and puts back what will actually be used. Anything that is
    /// not a number, or is outside a minute to a day, is replaced rather than argued with —
    /// there is nothing here worth stopping the user over.
    /// </summary>
    private void OnRefreshMinutesChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        if (int.TryParse(RefreshMinutes.Text, out var typed) && typed > 0)
        {
            Settings.OutlookCalendarRefreshMinutes = typed;
        }

        var minutes = Settings.EffectiveOutlookRefreshMinutes;
        Settings.OutlookCalendarRefreshMinutes = minutes;
        RefreshMinutes.Text = minutes.ToString(Ko.Culture);

        _loading = true;
        Refresh5.IsChecked = minutes == 5;
        Refresh10.IsChecked = minutes == 10;
        Refresh30.IsChecked = minutes == 30;
        Refresh60.IsChecked = minutes == 60;
        RefreshCustom.IsChecked = minutes is not (5 or 10 or 30 or 60);
        _loading = false;

        DescribeRefresh();

        _shell.ApplyCalendarOverlay();
        _shell.SaveSettings();
    }

    private void OnParserOptionChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.AssumeAfternoonForBareHours = AfternoonBias.IsChecked == true;
        _shell.ApplyParserSettings();
        _shell.SaveSettings();
    }

    /// <summary>
    /// Says what will actually happen at the next sign-in, which is not always what the
    /// checkbox says. An entry can be registered and still never run — pointing at a build
    /// that has moved, or switched off in Task Manager, where Windows quietly wins.
    /// </summary>
    private void DescribeStartup()
    {
        StartupHint.Text = StartupService.Current() switch
        {
            StartupState.BlockedByWindows =>
                "Windows에서 이 항목이 꺼져 있습니다. 체크를 껐다 켜면 다시 켜집니다.",

            StartupState.StalePath =>
                "예전 위치의 실행 파일을 가리키고 있습니다. 체크를 껐다 켜면 지금 위치로 등록됩니다.",

            StartupState.On => "로그인하면 트레이에 뜹니다",
            _ => "로그인할 때 자동으로 실행합니다",
        };
    }

    private void OnStartupChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        var wanted = LaunchAtStartup.IsChecked == true;
        if (StartupService.Set(wanted))
        {
            Settings.LaunchAtStartup = wanted;
            _shell.SaveSettings();
            DescribeStartup();
            Status(wanted ? "시작 프로그램에 등록했습니다" : "시작 프로그램에서 제거했습니다");
            return;
        }

        // The registry refused; put the checkbox back so it does not lie.
        _loading = true;
        LaunchAtStartup.IsChecked = !wanted;
        _loading = false;
        Status("시작 프로그램 설정을 변경하지 못했습니다");
    }

    private void OnShowOnStartChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.ShowWidgetOnStart = ShowWidgetOnStart.IsChecked == true;
        _shell.SaveSettings();
    }

    private void OnWatchInboxChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.EnableInboxWatch = WatchInbox.IsChecked == true;
        _shell.SaveSettings();
        _shell.ApplyInboxWatch();
        DescribeInbox();
    }

    /// <summary>
    /// Names the folder outright. Somebody setting up the sending application has to type
    /// this path in over there, and a description that does not include it is no help.
    /// </summary>
    private void DescribeInbox() =>
        InboxHint.Text = WatchInbox.IsChecked == true
            ? $"{_shell.InboxFolder} 에 들어온 Flowdeck 파일을 자동으로 가져옵니다"
            : "받는 폴더에 파일이 들어와도 가져오지 않습니다";

    private void OnOpenInboxClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Directory.CreateDirectory(_shell.InboxFolder);
            Process.Start(new ProcessStartInfo(_shell.InboxFolder) { UseShellExecute = true });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or Win32Exception)
        {
            Status("받는 폴더를 열지 못했습니다: " + ex.Message);
        }
    }

    /// <summary>Turns the pressed combination into a gesture string, if it is a usable one.</summary>
    private void OnHotkeyCapture(object sender, KeyEventArgs e)
    {
        e.Handled = true;

        var key = e.Key == Key.System ? e.SystemKey : e.Key;

        if (key is Key.LeftCtrl or Key.RightCtrl or Key.LeftAlt or Key.RightAlt
            or Key.LeftShift or Key.RightShift or Key.LWin or Key.RWin)
        {
            return;
        }

        var parts = new List<string>();
        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0) parts.Add("Ctrl");
        if ((Keyboard.Modifiers & ModifierKeys.Alt) != 0) parts.Add("Alt");
        if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0) parts.Add("Shift");
        if ((Keyboard.Modifiers & ModifierKeys.Windows) != 0) parts.Add("Win");

        if (parts.Count == 0)
        {
            Warn("수정 키(Ctrl, Alt, Shift, Win)를 하나 이상 포함해야 합니다");
            return;
        }

        parts.Add(key.ToString());
        var gesture = string.Join("+", parts);

        if (ReferenceEquals(sender, QuickAddHotkey))
        {
            QuickAddHotkey.Text = gesture;
            Settings.QuickAddHotkey = gesture;
        }
        else if (ReferenceEquals(sender, ToggleWidgetHotkey))
        {
            ToggleWidgetHotkey.Text = gesture;
            Settings.ToggleWidgetHotkey = gesture;
        }
        else if (ReferenceEquals(sender, AgendaEventsHotkey))
        {
            AgendaEventsHotkey.Text = gesture;
            Settings.AgendaEventsHotkey = gesture;
        }
        else if (ReferenceEquals(sender, AgendaTodosHotkey))
        {
            AgendaTodosHotkey.Text = gesture;
            Settings.AgendaTodosHotkey = gesture;
        }
        else if (ReferenceEquals(sender, CaptureSelectionHotkey))
        {
            CaptureSelectionHotkey.Text = gesture;
            Settings.CaptureSelectionHotkey = gesture;
        }
        else
        {
            return;
        }

        var problem = _shell.ReapplyHotKeys();
        Warn(problem);
        _shell.SaveSettings();
    }

    /// <summary>
    /// Opens the address book. The editor writes each change through as it is made, so
    /// there is nothing to collect when it closes — only the summary line to refresh.
    /// </summary>
    private void OnContactsClick(object sender, RoutedEventArgs e)
    {
        var editor = new ContactsWindow(new ContactsViewModel(Settings.Contacts, _shell.SaveSettings))
        {
            Owner = this,
        };

        editor.ShowDialog();

        _shell.ApplyParserSettings();
        DescribeContacts();
    }

    private void DescribeContacts()
    {
        var count = Settings.Contacts.Aliases.Count;
        ContactsSummary.Text = count == 0
            ? "@이름 으로 참석자를 부르려면 여기에 등록하세요"
            : $"{count}명 등록됨. 입력할 때 @이름 으로 부를 수 있습니다";
    }

    /// <summary>
    /// Folds a file exported elsewhere into this workspace. Additive: nothing already here
    /// is replaced or removed, so the worst a wrong file can do is add entries the user can
    /// then delete.
    /// </summary>
    private async void OnImportClick(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            DefaultExt = TransferFile.Extension,
            Filter = "Flowdeck 내보내기 (*.txt;*.json)|*.txt;*.json|모든 파일 (*.*)|*.*",
        };

        if (dialog.ShowDialog(this) != true) return;

        try
        {
            var archive = TransferFile.Read(File.ReadAllText(dialog.FileName));
            var result = await _shell.Repository.ImportAsync(archive);
            Status(result.Describe());
        }
        catch (FormatException ex)
        {
            Status(ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Status("파일을 열지 못했습니다: " + ex.Message);
        }
    }

    private void OnExportClick(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"flowdeck-{DateTime.Now:yyyyMMdd}.ics",
            DefaultExt = ".ics",
            Filter = "iCalendar (*.ics)|*.ics",
        };

        if (dialog.ShowDialog(this) != true) return;

        try
        {
            var ics = IcsExporter.Export(_shell.Repository.Events, DateTime.Now);
            File.WriteAllText(dialog.FileName, ics);
            Status($"{_shell.Repository.Events.Count}건의 일정을 내보냈습니다");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Status("파일을 저장하지 못했습니다: " + ex.Message);
        }
    }

    private void OnOpenFolderClick(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(_shell.DataFolder);
        Process.Start(new ProcessStartInfo(_shell.DataFolder) { UseShellExecute = true });
    }

    private void OnHeaderDrag(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left) return;

        try
        {
            DragMove();
        }
        catch (InvalidOperationException)
        {
        }
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    private void Status(string message) => StatusText.Text = message;

    private void Warn(string message)
    {
        HotkeyWarning.Text = message;
        HotkeyWarning.Visibility = string.IsNullOrEmpty(message) ? Visibility.Collapsed : Visibility.Visible;
    }
}

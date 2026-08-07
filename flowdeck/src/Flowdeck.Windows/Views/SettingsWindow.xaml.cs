using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Flowdeck.Core.Export;
using Flowdeck.Core.Settings;
using Flowdeck.Windows.Interop;

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

        KeywordHints.IsChecked = Settings.Routing.UseKeywordHints;
        AfternoonBias.IsChecked = Settings.AssumeAfternoonForBareHours;

        LaunchAtStartup.IsChecked = StartupService.IsEnabled();
        ShowWidgetOnStart.IsChecked = Settings.ShowWidgetOnStart;

        _loading = false;

        // These two have no Checked handler of their own; they commit on close.
        KeywordHints.Checked += OnParserOptionChanged;
        KeywordHints.Unchecked += OnParserOptionChanged;
        AfternoonBias.Checked += OnParserOptionChanged;
        AfternoonBias.Unchecked += OnParserOptionChanged;
        ColorWeekends.Checked += OnWeekStartChanged;
        ColorWeekends.Unchecked += OnWeekStartChanged;
        LaunchAtStartup.Checked += OnStartupChanged;
        LaunchAtStartup.Unchecked += OnStartupChanged;
        ShowWidgetOnStart.Checked += OnShowOnStartChanged;
        ShowWidgetOnStart.Unchecked += OnShowOnStartChanged;
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

    private void OnParserOptionChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        Settings.Routing.UseKeywordHints = KeywordHints.IsChecked == true;
        Settings.AssumeAfternoonForBareHours = AfternoonBias.IsChecked == true;
        _shell.ApplyParserSettings();
        _shell.SaveSettings();
    }

    private void OnStartupChanged(object sender, RoutedEventArgs e)
    {
        if (_loading) return;

        var wanted = LaunchAtStartup.IsChecked == true;
        if (StartupService.Set(wanted))
        {
            Settings.LaunchAtStartup = wanted;
            _shell.SaveSettings();
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
        else
        {
            return;
        }

        var problem = _shell.ReapplyHotKeys();
        Warn(problem);
        _shell.SaveSettings();
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

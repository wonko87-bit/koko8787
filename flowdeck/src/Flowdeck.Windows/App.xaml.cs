using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Threading;
using System.Windows;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Settings;
using Flowdeck.Core.Storage;
using Flowdeck.Windows.Integration;
using Flowdeck.Windows.Interop;
using Flowdeck.Windows.Services;
using Flowdeck.Windows.ViewModels;
using Flowdeck.Windows.Views;

namespace Flowdeck.Windows;

/// <summary>
/// Composition root. Flowdeck lives in the notification area rather than in a main
/// window, so the application object owns the widget, the capture overlay, the hot
/// keys and the reminder loop, and outlives all of them.
/// </summary>
public partial class App : Application, IAppShell
{
    private const string InstanceMutexName = @"Local\Flowdeck.SingleInstance";

    private Mutex? _instanceMutex;
    private HotKeyService? _hotKeys;
    private TrayIconController? _tray;
    private ReminderService? _reminders;

    private WorkspaceRepository? _repository;
    private NaturalLanguageParser? _parser;
    private ExternalCalendarFeed? _calendar;
    private DispatcherTimer? _calendarTimer;
    private AppSettings _settings = new();

    private WidgetWindow? _widget;
    private QuickAddWindow? _quickAdd;
    private SettingsWindow? _settingsWindow;
    private OutlookBridge? _outlook;
    private AgendaWindow? _eventsAgenda;
    private AgendaWindow? _todosAgenda;

    public AppSettings Settings => _settings;

    public WorkspaceRepository Repository =>
        _repository ?? throw new InvalidOperationException("Startup has not finished.");

    public string DataFolder { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Flowdeck");

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // A second copy would fight the first over the hot keys and the workspace file.
        _instanceMutex = new Mutex(initiallyOwned: true, InstanceMutexName, out var isFirstInstance);
        if (!isFirstInstance)
        {
            Shutdown();
            return;
        }

        // No main window: closing the widget must not end the process.
        ShutdownMode = ShutdownMode.OnExplicitShutdown;

        // Installed before anything else so even a failure during startup is written down.
        Directory.CreateDirectory(DataFolder);
        CrashReporter.Install(DataFolder, (title, message) => _tray?.Notify(title, message));

        try
        {
            await StartAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Flowdeck를 시작하지 못했습니다.\n\n" + ex.Message,
                "Flowdeck",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown();
        }
    }

    private async Task StartAsync()
    {
        Directory.CreateDirectory(DataFolder);

        _settings = AppSettings.LoadFrom(Path.Combine(DataFolder, "settings.json"));
        ThemeManager.Apply(_settings.Theme);

        var store = new JsonWorkspaceStore(Path.Combine(DataFolder, "workspace.json"));

        // Push-only: Flowdeck writes into Outlook and never reads back, so there is
        // nothing to reconcile. Reading is a later addition behind the same interface.
        _outlook = new OutlookBridge();
        _repository = new WorkspaceRepository(store, _outlook) { ExternalEnabled = _settings.EnableOutlook };
        await _repository.LoadAsync();

        _parser = new NaturalLanguageParser(_settings.Routing);
        ApplyParserSettings();

        _calendar = new ExternalCalendarFeed(_outlook) { IsEnabled = _settings.EnableOutlook && _settings.ShowOutlookCalendar };

        var widgetViewModel = new WidgetViewModel(_repository, _parser, clock: null, calendar: _calendar);
        widgetViewModel.Captured += OnCaptured;

        _widget = new WidgetWindow(widgetViewModel, _settings);
        _widget.SettingsRequested += (_, _) => ShowSettings();
        _widget.HideRequested += (_, _) => _tray?.SetWidgetVisible(false);
        _widget.EventsAgendaRequested += (_, _) => ToggleEventsAgenda();
        _widget.TodosAgendaRequested += (_, _) => ToggleTodosAgenda();

        var quickAddViewModel = new QuickAddViewModel(_repository, _parser);
        quickAddViewModel.Captured += OnCaptured;
        _quickAdd = new QuickAddWindow(quickAddViewModel);

        _eventsAgenda = new AgendaWindow(
            new AgendaViewModel(_repository, AgendaMode.Events),
            _settings.AgendaEventsPlacement,
            _settings);

        _todosAgenda = new AgendaWindow(
            new AgendaViewModel(_repository, AgendaMode.Todos),
            _settings.AgendaTodosPlacement,
            _settings);

        _tray = new TrayIconController();
        _tray.QuickAddRequested += (_, _) => ShowQuickAdd();
        _tray.WidgetToggleRequested += (_, _) => ToggleWidget();
        _tray.EventsAgendaRequested += (_, _) => ToggleEventsAgenda();
        _tray.TodosAgendaRequested += (_, _) => ToggleTodosAgenda();
        _tray.SettingsRequested += (_, _) => ShowSettings();
        _tray.ExitRequested += (_, _) => Shutdown();

        _hotKeys = new HotKeyService();
        var hotKeyProblem = ReapplyHotKeys();

        // A swipe-delete has no window to report into, so the repository says it here.
        _repository.ExternalWarning += (_, message) => Dispatcher.Invoke(
            () => _tray?.Notify(_repository.External?.DisplayName ?? "Outlook", message));

        _reminders = new ReminderService(_repository, (title, message) => _tray?.Notify(title, message));
        _reminders.Start();

        // A new build in a new folder leaves the old registration pointing at nothing, and
        // the app would silently stop starting with Windows. Put it back where it belongs.
        if (_settings.LaunchAtStartup) StartupService.RepairIfStale();

        ApplyCalendarOverlay();

        if (_settings.ShowWidgetOnStart) ShowWidget();
        _tray.SetWidgetVisible(_widget.IsVisible);

        if (hotKeyProblem.Length > 0) _tray.Notify("Flowdeck", hotKeyProblem);
    }

    /// <summary>
    /// Runs after a line has been filed. The entry is saved either way; this reports a
    /// copy to Outlook that did not happen, and opens the Teams form when one was asked for.
    /// </summary>
    private void OnCaptured(object? sender, CaptureResult result)
    {
        if (result.ExternalError is not null)
        {
            _tray?.Notify("Outlook 저장 실패", result.ExternalError + " (항목은 저장되었습니다)");
        }

        if (result.LaunchUrl is not null) OpenExternally(result.LaunchUrl);
    }

    /// <summary>
    /// Hands a link to the shell, which passes it to Teams or to the browser. Failing to
    /// open one is not worth taking the app down for: the entry is already saved, and the
    /// notification says what did not happen.
    /// </summary>
    private void OpenExternally(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            _tray?.Notify("Teams를 열지 못했습니다", "항목은 저장되었습니다");
        }
    }

    // ---- IAppShell ---------------------------------------------------------

    public void ApplyTheme(AppTheme theme) => ThemeManager.Apply(theme);

    public void ApplyWidgetSettings() => _widget?.ApplySettings();

    /// <summary>
    /// Turns the read-only Outlook overlay on or off and puts its timer on the interval the
    /// user chose. Reading a mailbox is not something to be doing quietly in the background
    /// when nobody asked, so switching it off stops the timer and drops what was read.
    /// </summary>
    public void ApplyCalendarOverlay()
    {
        if (_calendar is null) return;

        _calendarTimer?.Stop();
        _calendar.IsEnabled = _settings.EnableOutlook && _settings.ShowOutlookCalendar;

        if (!_calendar.IsEnabled || !_calendar.IsAvailable)
        {
            _calendar.Clear();
            return;
        }

        _calendarTimer ??= new DispatcherTimer();
        _calendarTimer.Interval = TimeSpan.FromMinutes(_settings.EffectiveOutlookRefreshMinutes);
        _calendarTimer.Tick -= OnCalendarTick;
        _calendarTimer.Tick += OnCalendarTick;
        _calendarTimer.Start();

        OnCalendarTick(this, EventArgs.Empty);
    }

    /// <summary>
    /// Re-reads a window around today. Wide enough that paging a month either way still has
    /// something to show, and bounded so a diary full of never-ending repeats cannot turn
    /// one read into an unbounded walk.
    /// </summary>
    private void OnCalendarTick(object? sender, EventArgs e)
    {
        if (_calendar is null || !_calendar.IsEnabled) return;

        var today = DateTime.Now.Date;
        CrashReporter.Observe(
            _calendar.RefreshAsync(today.AddDays(-45), today.AddDays(120)),
            "ExternalCalendarFeed.Refresh");
    }

    public void ApplyParserSettings()
    {
        if (_parser is null) return;

        _parser.AssumeAfternoonForBareHours = _settings.AssumeAfternoonForBareHours;
        _parser.FirstDayOfWeek = _settings.FirstDayOfWeek;
        _parser.PushExternalByDefault = _settings.PushToOutlookByDefault;
        _parser.Contacts = _settings.Contacts;
    }

    public string ReapplyHotKeys()
    {
        var hotKeys = _hotKeys;
        if (hotKeys is null) return string.Empty;

        hotKeys.UnregisterAll();

        var problems = new StringBuilder();
        Claim(_settings.QuickAddHotkey, "빠른 입력", ShowQuickAdd);
        Claim(_settings.ToggleWidgetHotkey, "위젯 표시/숨기기", ToggleWidget);
        Claim(_settings.AgendaEventsHotkey, "일정 목록", ToggleEventsAgenda);
        Claim(_settings.AgendaTodosHotkey, "할일 목록", ToggleTodosAgenda);
        return problems.ToString().TrimEnd();

        void Claim(string gesture, string label, Action action)
        {
            switch (hotKeys.Register(gesture, action))
            {
                case HotKeyResult.AlreadyTaken:
                    // Either another application owns it, or it duplicates one of ours.
                    problems.AppendLine($"{label} 단축키 {gesture} 은(는) 이미 사용 중인 조합입니다.");
                    break;

                case HotKeyResult.InvalidGesture:
                    problems.AppendLine($"{label} 단축키 {gesture} 을(를) 인식하지 못했습니다.");
                    break;
            }
        }
    }

    public void SaveSettings()
    {
        try
        {
            _settings.SaveTo(Path.Combine(DataFolder, "settings.json"));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Losing a preference is not worth interrupting the user over.
        }
    }

    public void ShowWidget()
    {
        if (_widget is null) return;

        _widget.Show();
        _widget.ApplySettings();
        _tray?.SetWidgetVisible(true);
    }

    public void HideWidget()
    {
        _widget?.Hide();
        _tray?.SetWidgetVisible(false);
    }

    public void ShowQuickAdd() => _quickAdd?.Summon();

    public void ToggleEventsAgenda() => _eventsAgenda?.Toggle();

    public void ToggleTodosAgenda() => _todosAgenda?.Toggle();

    // ---- internals ---------------------------------------------------------

    private void ToggleWidget()
    {
        if (_widget is null) return;

        if (_widget.IsVisible)
        {
            HideWidget();
        }
        else
        {
            ShowWidget();
        }
    }

    private void ShowSettings()
    {
        if (_settingsWindow is { IsVisible: true })
        {
            _settingsWindow.Activate();
            return;
        }

        _settingsWindow = new SettingsWindow(this);
        _settingsWindow.Closed += (_, _) => _settingsWindow = null;
        _settingsWindow.Show();
        _settingsWindow.Activate();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // The widget records its position as it moves; make sure the last one sticks.
        if (_widget is not null) SaveSettings();

        _calendarTimer?.Stop();
        _reminders?.Dispose();
        _hotKeys?.Dispose();
        _tray?.Dispose();
        _instanceMutex?.Dispose();

        base.OnExit(e);
    }
}

using System.IO;
using System.Text;
using System.Windows;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Settings;
using Flowdeck.Core.Storage;
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
    private AppSettings _settings = new();

    private WidgetWindow? _widget;
    private QuickAddWindow? _quickAdd;
    private SettingsWindow? _settingsWindow;
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
        _repository = new WorkspaceRepository(store);
        await _repository.LoadAsync();

        _parser = new NaturalLanguageParser(_settings.Routing);
        ApplyParserSettings();

        _widget = new WidgetWindow(new WidgetViewModel(_repository, _parser), _settings);
        _widget.SettingsRequested += (_, _) => ShowSettings();
        _widget.HideRequested += (_, _) => _tray?.SetWidgetVisible(false);
        _widget.EventsAgendaRequested += (_, _) => ToggleEventsAgenda();
        _widget.TodosAgendaRequested += (_, _) => ToggleTodosAgenda();

        _quickAdd = new QuickAddWindow(new QuickAddViewModel(_repository, _parser));

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

        _reminders = new ReminderService(_repository, (title, message) => _tray?.Notify(title, message));
        _reminders.Start();

        if (_settings.ShowWidgetOnStart) ShowWidget();
        _tray.SetWidgetVisible(_widget.IsVisible);

        if (hotKeyProblem.Length > 0) _tray.Notify("Flowdeck", hotKeyProblem);
    }

    // ---- IAppShell ---------------------------------------------------------

    public void ApplyTheme(AppTheme theme) => ThemeManager.Apply(theme);

    public void ApplyWidgetSettings() => _widget?.ApplySettings();

    public void ApplyParserSettings()
    {
        if (_parser is null) return;

        _parser.AssumeAfternoonForBareHours = _settings.AssumeAfternoonForBareHours;
        _parser.FirstDayOfWeek = _settings.FirstDayOfWeek;
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

        _reminders?.Dispose();
        _hotKeys?.Dispose();
        _tray?.Dispose();
        _instanceMutex?.Dispose();

        base.OnExit(e);
    }
}

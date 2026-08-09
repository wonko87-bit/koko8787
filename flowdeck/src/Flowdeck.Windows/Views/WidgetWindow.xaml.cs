using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Threading;
using Flowdeck.Core.Settings;
using Flowdeck.Windows.Interop;
using Flowdeck.Windows.ViewModels;

namespace Flowdeck.Windows.Views;

/// <summary>
/// The always-available desktop panel: month grid, the selected day's events and
/// todos, and an inline capture box.
/// </summary>
public partial class WidgetWindow : Window
{
    /// <summary>
    /// Width bounds. Width alone sets the month grid's proportions, so it stays tightly
    /// bounded — the day cells are square and would look wrong at any other scale.
    /// </summary>
    private const double MinWidgetWidth = 300;
    private const double MaxWidgetWidth = 440;

    /// <summary>
    /// Height is free within these bounds and does not affect the calendar at all.
    /// Every extra pixel goes to the events and todos list, which is the only row that
    /// stretches, so a taller widget simply shows more of them.
    /// </summary>
    private const double MaxWidgetHeight = 1000;

    /// <summary>
    /// Everything the height needs besides the month grid and the list: margins, the
    /// header, the weekday row, the divider and the capture box.
    /// </summary>
    private const double VerticalChrome = 175;

    /// <summary>The least the list may be squeezed to — roughly three rows.</summary>
    private const double MinListHeight = 90;

    /// <summary>Left and right margins plus the border, subtracted to get the grid's width.</summary>
    private const double HorizontalChrome = 30;

    private readonly AppSettings _settings;
    private readonly WindowPinService _pin;
    private readonly DispatcherTimer _clockTimer;

    public WidgetWindow(WidgetViewModel viewModel, AppSettings settings)
    {
        InitializeComponent();

        _settings = settings;
        DataContext = viewModel;
        ViewModel = viewModel;

        _pin = new WindowPinService(this);

        ApplyGeometry();
        ApplySettings();

        // Catches midnight, and any overdue todo tipping over as the hour turns.
        _clockTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
        _clockTimer.Tick += (_, _) => ViewModel.TickClock();
        _clockTimer.Start();

        Closing += OnClosing;
    }

    public WidgetViewModel ViewModel { get; }

    /// <summary>Raised when the user dismisses the widget, so the tray menu can stay in step.</summary>
    public event EventHandler? HideRequested;

    /// <summary>Raised when the settings button is pressed.</summary>
    public event EventHandler? SettingsRequested;

    /// <summary>Raised by the "전체" shortcuts beside each section heading.</summary>
    public event EventHandler? EventsAgendaRequested;

    public event EventHandler? TodosAgendaRequested;

    /// <summary>Re-reads opacity and pin mode after the settings window changes them.</summary>
    public void ApplySettings()
    {
        Root.Opacity = Math.Clamp(_settings.WidgetOpacity, 0.35, 1.0);
        _pin.Apply(_settings.PinMode);
        ViewModel.ApplyCalendarSettings(_settings);
    }

    /// <summary>Puts the caret in the inline box, for the "capture here" hot key.</summary>
    public void FocusInput()
    {
        Activate();
        InlineInput.Focus();
        InlineInput.CaretIndex = InlineInput.Text.Length;
    }

    /// <summary>
    /// The shortest the widget may be at its current width. Grows with the calendar,
    /// since the month grid is sized from the width and cannot be cropped.
    /// </summary>
    private double MinimumHeight => VerticalChrome + (ViewModel.DayCellSize * 6) + MinListHeight;

    private void ApplyGeometry()
    {
        Width = Math.Clamp(_settings.WidgetWidth, MinWidgetWidth, MaxWidgetWidth);
        ViewModel.UpdateMetrics(Width - HorizontalChrome);
        ApplyHeight(_settings.WidgetHeight);

        var area = SystemParameters.WorkArea;

        // No stored position means first run: park the widget at the top right.
        var left = _settings.WidgetLeft ?? area.Right - Width - 24;
        var top = _settings.WidgetTop ?? area.Top + 24;

        // Keep it reachable if a monitor was unplugged since the position was saved.
        Left = Math.Clamp(left, area.Left - Width + 80, area.Right - 80);
        Top = Math.Clamp(top, area.Top, area.Bottom - 60);
    }

    private void OnHeaderDrag(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left) return;

        // DragMove throws if the button was released between the event and the call.
        try
        {
            DragMove();
        }
        catch (InvalidOperationException)
        {
        }

        PersistGeometry();
    }

    /// <summary>
    /// Corner drag. The two axes move independently: width rescales the calendar within
    /// its narrow band, height just lengthens the list below it.
    /// </summary>
    private void OnResize(object sender, DragDeltaEventArgs e)
    {
        Width = Math.Clamp(Width + e.HorizontalChange, MinWidgetWidth, MaxWidgetWidth);
        ViewModel.UpdateMetrics(Width - HorizontalChrome);
        ApplyHeight(Height + e.VerticalChange);
        PersistGeometry();
    }

    private void OnWidgetSizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (!e.WidthChanged) return;

        ViewModel.UpdateMetrics(e.NewSize.Width - HorizontalChrome);

        // A wider widget has a taller calendar and so a taller floor. Raising MinHeight
        // lets WPF stretch the window itself if it has just become too short.
        MinHeight = MinimumHeight;
    }

    private void ApplyHeight(double desired)
    {
        var minimum = MinimumHeight;

        MinHeight = minimum;
        Height = Math.Clamp(desired, minimum, Math.Max(minimum, MaxWidgetHeight));
    }

    private void OnInlineInputKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Enter:
                e.Handled = true;
                if (ViewModel.SubmitCommand.CanExecute(null)) ViewModel.SubmitCommand.Execute(null);
                break;

            case Key.Escape:
                e.Handled = true;
                ViewModel.QuickInput = string.Empty;
                Keyboard.ClearFocus();
                break;
        }
    }

    private void OnEventsAgendaClick(object sender, RoutedEventArgs e) =>
        EventsAgendaRequested?.Invoke(this, EventArgs.Empty);

    private void OnTodosAgendaClick(object sender, RoutedEventArgs e) =>
        TodosAgendaRequested?.Invoke(this, EventArgs.Empty);

    private void OnSettingsClick(object sender, RoutedEventArgs e) =>
        SettingsRequested?.Invoke(this, EventArgs.Empty);

    private void OnHideClick(object sender, RoutedEventArgs e)
    {
        Hide();
        HideRequested?.Invoke(this, EventArgs.Empty);
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        PersistGeometry();
        _clockTimer.Stop();
        _pin.Dispose();
    }

    private void PersistGeometry()
    {
        _settings.WidgetLeft = Left;
        _settings.WidgetTop = Top;
        _settings.WidgetWidth = Width;
        _settings.WidgetHeight = Height;
    }
}

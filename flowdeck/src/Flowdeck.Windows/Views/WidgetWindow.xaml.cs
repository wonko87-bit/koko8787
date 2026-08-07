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

    /// <summary>Re-reads opacity and pin mode after the settings window changes them.</summary>
    public void ApplySettings()
    {
        Root.Opacity = Math.Clamp(_settings.WidgetOpacity, 0.35, 1.0);
        _pin.Apply(_settings.PinMode);
    }

    /// <summary>Puts the caret in the inline box, for the "capture here" hot key.</summary>
    public void FocusInput()
    {
        Activate();
        InlineInput.Focus();
        InlineInput.CaretIndex = InlineInput.Text.Length;
    }

    private void ApplyGeometry()
    {
        Width = Math.Max(MinWidth, _settings.WidgetWidth);
        Height = Math.Max(MinHeight, _settings.WidgetHeight);

        var area = SystemParameters.WorkArea;

        // NaN means "first run": park the widget at the top right of the work area.
        var left = double.IsNaN(_settings.WidgetLeft) ? area.Right - Width - 24 : _settings.WidgetLeft;
        var top = double.IsNaN(_settings.WidgetTop) ? area.Top + 24 : _settings.WidgetTop;

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

    private void OnResize(object sender, DragDeltaEventArgs e)
    {
        Width = Math.Max(MinWidth, Width + e.HorizontalChange);
        Height = Math.Max(MinHeight, Height + e.VerticalChange);
        PersistGeometry();
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

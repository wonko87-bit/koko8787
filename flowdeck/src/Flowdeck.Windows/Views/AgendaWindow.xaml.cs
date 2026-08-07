using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Flowdeck.Core.Settings;
using Flowdeck.Windows.ViewModels;

namespace Flowdeck.Windows.Views;

/// <summary>
/// The standalone list window. One instance holds the calendar, another the todos —
/// the mode lives in the view model, so both share this shell.
///
/// Like the quick-add overlay it is hidden rather than closed, so toggling it is
/// instant and its scroll position survives.
/// </summary>
public partial class AgendaWindow : Window
{
    private readonly WindowPlacement _placement;
    private readonly AppSettings _settings;

    public AgendaWindow(AgendaViewModel viewModel, WindowPlacement placement, AppSettings settings)
    {
        InitializeComponent();

        ViewModel = viewModel;
        DataContext = viewModel;
        _placement = placement;
        _settings = settings;

        if (viewModel.Mode == AgendaMode.Todos) viewModel.ShowCompleted = settings.AgendaShowCompleted;
        viewModel.MatchAllTags = settings.AgendaMatchAllTags;

        Title = "Flowdeck · " + viewModel.Title;
        ApplyPlacement();
    }

    public AgendaViewModel ViewModel { get; }

    /// <summary>Raised when the window is dismissed, so the tray menu can stay in step.</summary>
    public event EventHandler? Dismissed;

    /// <summary>Shows the window if it is hidden, hides it if it is already in front.</summary>
    public void Toggle()
    {
        if (IsVisible)
        {
            Dismiss();
            return;
        }

        // Start unfiltered every time. A filter left on from an earlier session would
        // silently hide entries, and there would be nothing on screen saying why.
        ViewModel.ClearTags();
        ViewModel.Reload();
        EnsureOnScreen();
        Show();
        Activate();
    }

    public void Dismiss()
    {
        PersistPlacement();
        Hide();
        Dismissed?.Invoke(this, EventArgs.Empty);
    }

    private void ApplyPlacement()
    {
        Width = Math.Max(MinWidth, _placement.Width);
        Height = Math.Max(MinHeight, _placement.Height);

        if (_placement.HasPosition)
        {
            Left = _placement.Left;
            Top = _placement.Top;
        }
        else
        {
            var area = SystemParameters.WorkArea;
            Left = area.Left + ((area.Width - Width) / 2);
            Top = area.Top + ((area.Height - Height) / 2);
        }

        EnsureOnScreen();
    }

    /// <summary>Pulls the window back into view if the monitor it was on has gone away.</summary>
    private void EnsureOnScreen()
    {
        var area = SystemParameters.WorkArea;
        Left = Math.Clamp(Left, area.Left - Width + 120, area.Right - 120);
        Top = Math.Clamp(Top, area.Top, area.Bottom - 80);
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

        PersistPlacement();
    }

    private void OnResize(object sender, DragDeltaEventArgs e)
    {
        Width = Math.Max(MinWidth, Width + e.HorizontalChange);
        Height = Math.Max(MinHeight, Height + e.VerticalChange);
        PersistPlacement();
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape) return;

        e.Handled = true;
        Dismiss();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Dismiss();

    private void OnMatchModeClick(object sender, RoutedEventArgs e)
    {
        ViewModel.MatchAllTags = !ViewModel.MatchAllTags;
        _settings.AgendaMatchAllTags = ViewModel.MatchAllTags;
    }

    private void PersistPlacement()
    {
        _placement.Left = Left;
        _placement.Top = Top;
        _placement.Width = Width;
        _placement.Height = Height;

        if (ViewModel.Mode == AgendaMode.Todos) _settings.AgendaShowCompleted = ViewModel.ShowCompleted;
    }
}

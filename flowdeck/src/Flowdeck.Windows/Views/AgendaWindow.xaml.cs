using System.IO;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Flowdeck.Core.Export;
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

        if (_placement.Left.HasValue && _placement.Top.HasValue)
        {
            Left = _placement.Left.Value;
            Top = _placement.Top.Value;
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

    /// <summary>
    /// Writes out exactly what the window is showing. The tag chips are the picker: whatever
    /// they have narrowed the list to is what lands in the file, so there is no second idea
    /// of "selected" that could disagree with what is on screen.
    /// </summary>
    private void OnExportClick(object sender, RoutedEventArgs e)
    {
        var (todos, events) = ViewModel.VisibleEntries();
        if (todos.Count + events.Count == 0)
        {
            MessageBox.Show(this, "내보낼 항목이 없습니다.", "Flowdeck", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"flowdeck-{ViewModel.Title}-{DateTime.Now:yyyyMMdd}{TransferFile.Extension}",
            DefaultExt = TransferFile.Extension,
            Filter = $"Flowdeck 내보내기 (*{TransferFile.Extension})|*{TransferFile.Extension}|모든 파일 (*.*)|*.*",
        };

        if (dialog.ShowDialog(this) != true) return;

        try
        {
            File.WriteAllText(dialog.FileName, TransferFile.Write(todos, events, DateTimeOffset.Now));
            MessageBox.Show(
                this,
                $"{todos.Count + events.Count}건을 내보냈습니다.",
                "Flowdeck",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            MessageBox.Show(this, "파일을 저장하지 못했습니다.\n\n" + ex.Message, "Flowdeck", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

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

using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Flowdeck.Core.Models;
using Flowdeck.Windows.ViewModels;

namespace Flowdeck.Windows.Views;

/// <summary>
/// The hot-key overlay. It is created once and hidden rather than closed, so
/// summoning it is instant and the caret is always ready.
/// </summary>
public partial class QuickAddWindow : Window
{
    private readonly QuickAddViewModel _viewModel;

    public QuickAddWindow(QuickAddViewModel viewModel)
    {
        InitializeComponent();

        _viewModel = viewModel;
        DataContext = viewModel;
    }

    /// <summary>Brings the window up centred near the top of the active screen, ready to type.</summary>
    public void Summon()
    {
        _viewModel.Reset();
        PositionOnActiveScreen();

        Show();
        Activate();
        Input.Focus();
        Keyboard.Focus(Input);
    }

    public void Dismiss()
    {
        _viewModel.Reset();
        Hide();
    }

    private void PositionOnActiveScreen()
    {
        // WorkArea follows the primary screen; the cursor tells us which screen the
        // user is actually looking at, so the overlay opens where their attention is.
        var screen = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position);
        var bounds = screen.WorkingArea;

        var dpi = VisualTreeHelper.GetDpi(this);
        var left = bounds.Left / dpi.DpiScaleX + ((bounds.Width / dpi.DpiScaleX) - Width) / 2;
        var top = bounds.Top / dpi.DpiScaleY + (bounds.Height / dpi.DpiScaleY) * 0.22;

        Left = left;
        Top = top;
    }

    private async void OnInputKeyDown(object sender, KeyEventArgs e)
    {
        var ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

        if (ctrl)
        {
            switch (e.Key)
            {
                case Key.D1 or Key.NumPad1:
                    e.Handled = true;
                    _viewModel.ForceTarget(EntryTarget.Todo);
                    return;

                case Key.D2 or Key.NumPad2:
                    e.Handled = true;
                    _viewModel.ForceTarget(EntryTarget.Calendar);
                    return;

                case Key.D3 or Key.NumPad3:
                    e.Handled = true;
                    _viewModel.ForceTarget(EntryTarget.Both);
                    return;
            }
        }

        switch (e.Key)
        {
            case Key.Enter:
                e.Handled = true;
                if (await _viewModel.SubmitAsync()) Dismiss();
                break;

            case Key.Escape:
                e.Handled = true;
                Dismiss();
                break;
        }
    }

    /// <summary>Clicking away is the same as pressing Escape.</summary>
    private void OnDeactivated(object? sender, EventArgs e)
    {
        if (IsVisible) Dismiss();
    }
}

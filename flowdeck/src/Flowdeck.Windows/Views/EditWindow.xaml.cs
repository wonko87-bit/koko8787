using System.Windows;
using System.Windows.Input;
using Flowdeck.Windows.Services;
using Flowdeck.Windows.ViewModels;

namespace Flowdeck.Windows.Views;

/// <summary>
/// One entry, opened up. Reached by double-clicking a row in the widget or in either list
/// window — the desktop idiom for opening a thing, and the one gesture that will not fire
/// by accident on a widget that already sits under the cursor all day.
/// </summary>
public partial class EditWindow : Window
{
    public EditWindow(EditViewModel viewModel)
    {
        InitializeComponent();

        DataContext = viewModel;
        viewModel.Finished += (_, _) => Close();

        // The title is what usually needs fixing, and it is already selected to be typed over.
        Loaded += (_, _) =>
        {
            TitleBox.Focus();
            TitleBox.SelectAll();
        };
    }

    /// <summary>
    /// Escape abandons the edit from any field. Ctrl+Enter saves from any field, including
    /// the note box, where a plain Enter has to stay a line break.
    /// </summary>
    private void OnFieldKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Escape:
                e.Handled = true;
                Close();
                break;

            case Key.Enter when Keyboard.Modifiers == ModifierKeys.Control:
                e.Handled = true;
                if (DataContext is EditViewModel viewModel)
                {
                    CrashReporter.Observe(viewModel.SaveAsync(), "EditWindow.Save");
                }

                break;
        }
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
}

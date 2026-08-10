using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Flowdeck.Windows.ViewModels;

namespace Flowdeck.Windows.Views;

/// <summary>
/// The address-book editor: the same list that lives in settings.json, without needing to
/// open settings.json. Every change is written through as it is made, so there is no save
/// button and nothing to lose by closing the window.
/// </summary>
public partial class ContactsWindow : Window
{
    public ContactsWindow(ContactsViewModel viewModel)
    {
        InitializeComponent();

        DataContext = viewModel;
        Loaded += (_, _) => NewName.Focus();
    }

    /// <summary>
    /// The rows commit when they lose focus, and closing the window is not something that
    /// makes a text box lose focus first. Without this, the last edit made before pressing
    /// the close button would be thrown away — which looks exactly like the editor not
    /// saving at all.
    /// </summary>
    protected override void OnClosing(CancelEventArgs e)
    {
        if (Keyboard.FocusedElement is TextBox box)
        {
            box.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        }

        base.OnClosing(e);
    }

    /// <summary>Enter from either field adds the row, so a name never needs the mouse.</summary>
    private void OnFieldKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Enter:
                e.Handled = true;
                if (DataContext is ContactsViewModel viewModel) viewModel.Add();
                break;

            case Key.Escape:
                e.Handled = true;
                Close();
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

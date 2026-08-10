using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.ViewModels;

namespace Flowdeck.Mobile.Views;

public partial class CalendarPage : ContentPage
{
    private CalendarViewModel? _viewModel;

    public CalendarPage() => InitializeComponent();

    protected override async void OnAppearing()
    {
        base.OnAppearing();

        if (_viewModel is null)
        {
            await Workspace.EnsureLoadedAsync();
            _viewModel = new CalendarViewModel();
            BindingContext = _viewModel;
            return;
        }

        // Coming back from the capture tab, and possibly across midnight.
        _viewModel.Refresh();
    }
}

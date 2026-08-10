using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.ViewModels;

namespace Flowdeck.Mobile.Views;

public partial class ListPage : ContentPage
{
    private ListViewModel? _viewModel;

    public ListPage() => InitializeComponent();

    protected override async void OnAppearing()
    {
        base.OnAppearing();

        if (_viewModel is null)
        {
            await Workspace.EnsureLoadedAsync();
            _viewModel = new ListViewModel();
            BindingContext = _viewModel;
        }

        // Coming back from the capture tab, what was just typed has to be here.
        _viewModel.Reload();
    }
}

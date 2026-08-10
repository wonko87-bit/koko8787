using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.ViewModels;

namespace Flowdeck.Mobile.Views;

public partial class SettingsPage : ContentPage
{
    private SettingsViewModel? _viewModel;

    public SettingsPage() => InitializeComponent();

    protected override async void OnAppearing()
    {
        base.OnAppearing();

        if (_viewModel is not null) return;

        await Workspace.EnsureLoadedAsync();
        _viewModel = new SettingsViewModel();
        BindingContext = _viewModel;
    }
}

using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.ViewModels;

namespace Flowdeck.Mobile.Views;

public partial class CapturePage : ContentPage
{
    private CaptureViewModel? _viewModel;

    public CapturePage() => InitializeComponent();

    /// <summary>
    /// The view model is built here rather than in the constructor because it reads the
    /// parser, and the parser is not configured until the workspace has been loaded off disk.
    /// </summary>
    protected override async void OnAppearing()
    {
        base.OnAppearing();

        if (_viewModel is null)
        {
            await Workspace.EnsureLoadedAsync();
            _viewModel = new CaptureViewModel();
            BindingContext = _viewModel;
        }

        InputBox.Focus();
    }
}

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
            _viewModel.EditRequested += OnEditRequested;
            BindingContext = _viewModel;
        }

        // Coming back from the capture tab, what was just typed has to be here.
        _viewModel.Reload();
    }

    /// <summary>
    /// Opens the detail sheet over this page. Modal rather than a pushed page: it is one
    /// entry being looked at, not a place navigated to, and it closes back to exactly here.
    /// </summary>
    private async void OnEditRequested(object? sender, (string Id, bool IsTodo) row)
    {
        var editor = EditViewModel.For(row.Id, row.IsTodo);
        if (editor is null) return;

        await Navigation.PushModalAsync(new EditPage(editor));
    }
}

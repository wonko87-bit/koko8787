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
            _viewModel.EditRequested += OnEditRequested;
            BindingContext = _viewModel;
            return;
        }

        // Coming back from the capture tab, and possibly across midnight.
        _viewModel.Refresh();
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

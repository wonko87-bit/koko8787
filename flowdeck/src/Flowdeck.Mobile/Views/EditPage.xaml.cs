using Flowdeck.Mobile.ViewModels;

namespace Flowdeck.Mobile.Views;

/// <summary>
/// The detail sheet, pushed over whichever list the entry was tapped in. Closes itself once
/// the view model reports the entry saved or deleted.
/// </summary>
public partial class EditPage : ContentPage
{
    private readonly EditViewModel _viewModel;

    public EditPage(EditViewModel viewModel)
    {
        InitializeComponent();

        _viewModel = viewModel;
        BindingContext = viewModel;

        viewModel.Finished += (_, _) => MainThread.BeginInvokeOnMainThread(Close);
    }

    private async void OnCancelClick(object? sender, EventArgs e) => await CloseAsync();

    /// <summary>
    /// Deleting is the one thing here that cannot be undone, so it asks first. Everything
    /// else on this screen can be typed back.
    /// </summary>
    private async void OnDeleteClick(object? sender, EventArgs e)
    {
        var confirmed = await DisplayAlert("삭제", "이 항목을 지울까요?", "삭제", "취소");
        if (!confirmed) return;

        await _viewModel.DeleteAsync();
    }

    private void Close() => _ = CloseAsync();

    private async Task CloseAsync()
    {
        try
        {
            if (Navigation.ModalStack.Count > 0) await Navigation.PopModalAsync();
        }
        catch (Exception)
        {
            // Already gone — closing twice is not worth a crash.
        }
    }
}

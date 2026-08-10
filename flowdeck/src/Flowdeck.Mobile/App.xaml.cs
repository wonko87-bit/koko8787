namespace Flowdeck.Mobile;

public partial class App : Application
{
    public App() => InitializeComponent();

    /// <summary>
    /// The window is created here rather than by assigning MainPage, which is deprecated
    /// from .NET 9 onwards. Both work today; only one of them will still be here later.
    /// </summary>
    protected override Window CreateWindow(IActivationState? activationState) =>
        new(new AppShell());
}

namespace Flowdeck.Mobile;

public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder.UseMauiApp<App>();

        // No fonts are registered on purpose: the styles name no font family, so the app
        // takes the system face. On a phone that is what everything else uses anyway, and
        // it keeps a few hundred kilobytes of .ttf out of the package.
        return builder.Build();
    }
}

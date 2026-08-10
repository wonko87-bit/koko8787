using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile;

public partial class AppShell : Shell
{
    public AppShell()
    {
        InitializeComponent();

        Loaded += (_, _) => GoToPendingRoute();
    }

    /// <summary>
    /// Lands on whichever tab the launcher asked for. Tapping the month on the home screen
    /// and arriving at the capture box would be a small betrayal of what was tapped.
    /// </summary>
    public static void GoToPendingRoute()
    {
        var route = Workspace.PendingRoute;
        Workspace.PendingRoute = null;
        if (route is null || Current is null) return;

        try
        {
            Current.GoToAsync("//" + route);
        }
        catch (Exception)
        {
            // Landing on the wrong tab is not worth a crash.
        }
    }
}

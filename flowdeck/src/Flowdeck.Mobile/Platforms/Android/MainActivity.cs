using Android.App;
using Android.Content.PM;
using Android.OS;
using Flowdeck.Mobile.Reminders;
using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.Widgets;

namespace Flowdeck.Mobile;

[Activity(
    Theme = "@style/Maui.SplashTheme",
    MainLauncher = true,
    LaunchMode = LaunchMode.SingleTop,
    ConfigurationChanges = ConfigChanges.ScreenSize
        | ConfigChanges.Orientation
        | ConfigChanges.UiMode
        | ConfigChanges.ScreenLayout
        | ConfigChanges.SmallestScreenSize
        | ConfigChanges.Density)]
public class MainActivity : MauiAppCompatActivity
{
    private const int NotificationPermissionRequest = 4801;

    protected override void OnCreate(Bundle? savedInstanceState)
    {
        // Read before the shell exists, so it is waiting when the shell asks.
        Workspace.PendingRoute = Intent?.GetStringExtra("route");

        base.OnCreate(savedInstanceState);

        ReminderScheduler.EnsureChannel(this);
        AskForNotificationPermission();

        // Anything that changes an entry — here, the popup, or an import — has to leave the
        // alarms and the home screen agreeing with it.
        Workspace.EntriesChanged = () =>
        {
            ReminderScheduler.Reschedule(this);
            FlowdeckWidgets.RefreshAll(this);
        };
    }

    /// <summary>The app was already running when the widget was tapped.</summary>
    protected override void OnNewIntent(Android.Content.Intent? intent)
    {
        base.OnNewIntent(intent);

        Workspace.PendingRoute = intent?.GetStringExtra("route");
        AppShell.GoToPendingRoute();
    }

    protected override void OnResume()
    {
        base.OnResume();

        // Catches what happened while the app was away: a reminder fired from an alarm, or
        // an entry added from the popup without the app ever coming up.
        Task.Run(async () =>
        {
            await Workspace.EnsureLoadedAsync();
            RunOnUiThread(() => ReminderScheduler.Reschedule(this));
        });
    }

    /// <summary>
    /// Android 13 and later will not show a notification until asked. Asked once, on the
    /// first run: refusing is a decision the system remembers, and pestering does not change it.
    /// </summary>
    private void AskForNotificationPermission()
    {
        if (!OperatingSystem.IsAndroidVersionAtLeast(33)) return;

        if (CheckSelfPermission(Android.Manifest.Permission.PostNotifications) == Permission.Granted) return;

        RequestPermissions(
            new[] { Android.Manifest.Permission.PostNotifications },
            NotificationPermissionRequest);
    }
}

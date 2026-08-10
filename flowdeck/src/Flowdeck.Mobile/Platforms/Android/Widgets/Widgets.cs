using Android.App;
using Android.Appwidget;
using Android.Content;
using Android.Content.Res;
using Android.Widget;
using Flowdeck.Mobile.Services;

namespace Flowdeck.Mobile.Widgets;

/// <summary>Redraws whatever the user has actually put on their home screen.</summary>
public static class FlowdeckWidgets
{
    public static void RefreshAll(Context context)
    {
        Refresh(context, typeof(CalendarWidget));
        Refresh(context, typeof(InputWidget));
    }

    private static void Refresh(Context context, Type provider)
    {
        var manager = AppWidgetManager.GetInstance(context);
        if (manager is null) return;

        var ids = manager.GetAppWidgetIds(new ComponentName(context, provider));
        if (ids is null || ids.Length == 0) return;

        var update = new Intent(context, provider)
            .SetAction(AppWidgetManager.ActionAppwidgetUpdate)
            .PutExtra(AppWidgetManager.ExtraAppwidgetIds, ids);

        context.SendBroadcast(update);
    }

    internal static bool IsDark(Context context) =>
        (context.Resources?.Configuration?.UiMode & UiMode.NightMask) == UiMode.NightYes;
}

/// <summary>
/// The month, five cells by five. Tapping it opens the app on the calendar tab — a widget
/// this size is for reading, and anything finer than "open it" is a mis-tap waiting to
/// happen at this scale.
/// </summary>
[BroadcastReceiver(Label = "Flowdeck 달력", Enabled = true, Exported = true)]
[IntentFilter(new[] { AppWidgetManager.ActionAppwidgetUpdate })]
[MetaData("android.appwidget.provider", Resource = "@xml/calendar_widget_info")]
public sealed class CalendarWidget : AppWidgetProvider
{
    public override void OnUpdate(Context? context, AppWidgetManager? appWidgetManager, int[]? appWidgetIds)
    {
        if (context is null || appWidgetManager is null || appWidgetIds is null) return;

        var pending = GoAsync();

        Task.Run(async () =>
        {
            try
            {
                await Workspace.EnsureLoadedAsync();

                var views = new RemoteViews(context.PackageName, Resource.Layout.calendar_widget);
                views.SetImageViewBitmap(
                    Resource.Id.calendar_image,
                    MonthBitmap.Render(DateTime.Now, FlowdeckWidgets.IsDark(context)));

                views.SetOnClickPendingIntent(Resource.Id.calendar_root, OpenApp(context, "calendar"));

                foreach (var id in appWidgetIds) appWidgetManager.UpdateAppWidget(id, views);
            }
            catch (Exception)
            {
                // A widget that fails to draw should leave the last picture up rather than
                // take the process down behind it.
            }
            finally
            {
                pending.Finish();
            }
        });
    }

    internal static PendingIntent OpenApp(Context context, string route)
    {
        var intent = new Intent(context, typeof(MainActivity))
            .SetAction(Intent.ActionMain)
            .AddFlags(ActivityFlags.NewTask | ActivityFlags.SingleTop)
            .PutExtra("route", route);

        return PendingIntent.GetActivity(
            context,
            route.GetHashCode() & 0x7FFFFFFF,
            intent,
            PendingIntentFlags.UpdateCurrent | PendingIntentFlags.Immutable)!;
    }
}

/// <summary>
/// A single strip that opens the capture popup — the phone's answer to Ctrl+Alt+Space.
///
/// It is not a text field. A widget cannot hold a keyboard, and a strip that looked
/// editable but was not would be worse than one that plainly is a button.
/// </summary>
[BroadcastReceiver(Label = "Flowdeck 빠른 입력", Enabled = true, Exported = true)]
[IntentFilter(new[] { AppWidgetManager.ActionAppwidgetUpdate })]
[MetaData("android.appwidget.provider", Resource = "@xml/input_widget_info")]
public sealed class InputWidget : AppWidgetProvider
{
    public override void OnUpdate(Context? context, AppWidgetManager? appWidgetManager, int[]? appWidgetIds)
    {
        if (context is null || appWidgetManager is null || appWidgetIds is null) return;

        var views = new RemoteViews(context.PackageName, Resource.Layout.input_widget);

        var capture = new Intent(context, typeof(QuickCaptureActivity))
            .AddFlags(ActivityFlags.NewTask | ActivityFlags.ClearTop);

        var pending = PendingIntent.GetActivity(
            context,
            0,
            capture,
            PendingIntentFlags.UpdateCurrent | PendingIntentFlags.Immutable);

        views.SetOnClickPendingIntent(Resource.Id.input_root, pending);

        foreach (var id in appWidgetIds) appWidgetManager.UpdateAppWidget(id, views);
    }
}

using Android.App;
using Android.Appwidget;
using Android.Content;
using Android.Content.Res;
using Android.OS;
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

        // ComponentName wants the Java class, not the CLR type behind it.
        var ids = manager.GetAppWidgetIds(new ComponentName(context, Java.Lang.Class.FromType(provider)));
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

                var dark = FlowdeckWidgets.IsDark(context);
                var open = OpenApp(context, "calendar");

                // One picture per widget, not one for all: two copies on the home screen can
                // be different sizes, and each is drawn at its own.
                foreach (var id in appWidgetIds)
                {
                    var (width, height) = SizeOf(context, appWidgetManager, id);

                    var views = new RemoteViews(context.PackageName, Resource.Layout.calendar_widget);
                    views.SetImageViewBitmap(Resource.Id.calendar_image, MonthBitmap.Render(width, height, DateTime.Now, dark));
                    views.SetOnClickPendingIntent(Resource.Id.calendar_root, open);

                    appWidgetManager.UpdateAppWidget(id, views);
                }
            }
            catch (Exception)
            {
                // A widget that fails to draw should leave the last picture up rather than
                // take the process down behind it.
            }
            finally
            {
                pending?.Finish();
            }
        });
    }

    /// <summary>
    /// The launcher resized the widget. Nothing about the picture is right for the new
    /// shape, so it is drawn again for it.
    /// </summary>
    public override void OnAppWidgetOptionsChanged(Context? context, AppWidgetManager? appWidgetManager, int appWidgetId, Bundle? newOptions)
    {
        base.OnAppWidgetOptionsChanged(context, appWidgetManager, appWidgetId, newOptions);
        OnUpdate(context, appWidgetManager, new[] { appWidgetId });
    }

    /// <summary>
    /// The widget's size in pixels, as the launcher reports it.
    ///
    /// The launcher gives two sizes, one for each way the screen can be held: in portrait
    /// the widget is at its narrowest and tallest, in landscape at its widest and shortest.
    /// Those are read off the options as min-width by max-height and max-width by min-height,
    /// which is how the platform documents them, and turned into pixels with the screen's
    /// density. A launcher that has not said (a widget that was only just placed) gets the
    /// square the provider asks for.
    /// </summary>
    private static (int Width, int Height) SizeOf(Context context, AppWidgetManager manager, int id)
    {
        var options = manager.GetAppWidgetOptions(id);
        var resources = context.Resources;
        var density = resources?.DisplayMetrics?.Density ?? 2f;
        // Spelled out: Android.Widget has an Orientation of its own, and both are in scope here.
        var landscape = resources?.Configuration?.Orientation == Android.Content.Res.Orientation.Landscape;

        var minWidth = options?.GetInt(AppWidgetManager.OptionAppwidgetMinWidth) ?? 0;
        var maxWidth = options?.GetInt(AppWidgetManager.OptionAppwidgetMaxWidth) ?? 0;
        var minHeight = options?.GetInt(AppWidgetManager.OptionAppwidgetMinHeight) ?? 0;
        var maxHeight = options?.GetInt(AppWidgetManager.OptionAppwidgetMaxHeight) ?? 0;

        var widthDp = landscape ? maxWidth : minWidth;
        var heightDp = landscape ? minHeight : maxHeight;
        if (widthDp <= 0 || heightDp <= 0)
        {
            widthDp = 320;
            heightDp = 320;
        }

        var width = (int)Math.Round(widthDp * density);
        var height = (int)Math.Round(heightDp * density);

        // Kept in proportion when capped, so the shape the launcher reported survives.
        var longest = Math.Max(width, height);
        if (longest > MonthBitmap.MaxSide)
        {
            var scale = MonthBitmap.MaxSide / (float)longest;
            width = (int)(width * scale);
            height = (int)(height * scale);
        }

        return (width, height);
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

using Android.App;
using Android.Content;
using Android.OS;
using AndroidX.Core.App;
using Flowdeck.Core.Services;
using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.Widgets;

namespace Flowdeck.Mobile.Reminders;

/// <summary>
/// Turns the reminders <see cref="ReminderPlanner"/> works out into Android alarms.
///
/// Alarms rather than a background loop, because a phone kills background loops and is
/// right to. One alarm is set per upcoming reminder, and each firing sets the next batch —
/// so the schedule stays correct without anything having to stay awake to keep it so.
/// </summary>
public static class ReminderScheduler
{
    public const string ChannelId = "flowdeck.reminders";

    internal const string ActionFire = "com.flowdeck.mobile.REMINDER";

    /// <summary>How far ahead alarms are held. Beyond this, the next firing takes over.</summary>
    private static readonly TimeSpan Horizon = TimeSpan.FromDays(3);

    /// <summary>How late a reminder may surface before it is dropped as noise.</summary>
    private static readonly TimeSpan Grace = TimeSpan.FromMinutes(20);

    private const int MaxAlarms = 20;

    private static readonly object Gate = new();
    private static List<int> _placed = new();

    /// <summary>
    /// Replaces every alarm with the current set. Cheap and idempotent, so it can be called
    /// after any change without working out what actually moved.
    /// </summary>
    public static void Reschedule(Context context)
    {
        var alarms = context.GetSystemService(Context.AlarmService) as AlarmManager;
        if (alarms is null) return;

        var now = DateTime.Now;
        var wanted = ReminderPlanner.Upcoming(Workspace.Repository, now, Horizon, MaxAlarms);

        lock (Gate)
        {
            foreach (var id in _placed) alarms.Cancel(FirePendingIntent(context, id));

            var placed = new List<int>(wanted.Count);
            foreach (var reminder in wanted)
            {
                var delay = reminder.At - now;
                if (delay < TimeSpan.Zero) continue;

                var at = Java.Lang.JavaSystem.CurrentTimeMillis() + (long)delay.TotalMilliseconds;
                Set(alarms, at, FirePendingIntent(context, reminder.AlarmId));
                placed.Add(reminder.AlarmId);
            }

            _placed = placed;
        }
    }

    /// <summary>
    /// Exact where the system still allows it, approximate where it does not. An exact
    /// alarm needs a permission that recent Android only grants on request, and a reminder
    /// a few minutes loose is worth far more than one that never arrives.
    /// </summary>
    private static void Set(AlarmManager alarms, long atMillis, PendingIntent intent)
    {
        if (OperatingSystem.IsAndroidVersionAtLeast(31) && !alarms.CanScheduleExactAlarms())
        {
            alarms.SetAndAllowWhileIdle(AlarmType.RtcWakeup, atMillis, intent);
            return;
        }

        if (OperatingSystem.IsAndroidVersionAtLeast(23))
        {
            alarms.SetExactAndAllowWhileIdle(AlarmType.RtcWakeup, atMillis, intent);
            return;
        }

        alarms.SetExact(AlarmType.RtcWakeup, atMillis, intent);
    }

    private static PendingIntent FirePendingIntent(Context context, int id)
    {
        var intent = new Intent(context, typeof(ReminderReceiver)).SetAction(ActionFire);

        return PendingIntent.GetBroadcast(
            context,
            id,
            intent,
            PendingIntentFlags.UpdateCurrent | PendingIntentFlags.Immutable)!;
    }

    /// <summary>Created once; Android ignores the call after the first.</summary>
    public static void EnsureChannel(Context context)
    {
        if (!OperatingSystem.IsAndroidVersionAtLeast(26)) return;

        if (context.GetSystemService(Context.NotificationService) is not NotificationManager manager) return;

        var channel = new NotificationChannel(ChannelId, "알림", NotificationImportance.High)
        {
            Description = "일정과 할일 알림",
        };

        manager.CreateNotificationChannel(channel);
    }

    /// <summary>Posts whatever came due in the window just passed.</summary>
    internal static void PostDue(Context context)
    {
        EnsureChannel(context);

        foreach (var reminder in ReminderPlanner.Overdue(Workspace.Repository, DateTime.Now, Grace))
        {
            var open = PendingIntent.GetActivity(
                context,
                reminder.AlarmId,
                new Intent(context, typeof(MainActivity)).AddFlags(ActivityFlags.SingleTop),
                PendingIntentFlags.UpdateCurrent | PendingIntentFlags.Immutable);

            var builder = new NotificationCompat.Builder(context, ChannelId);
            builder.SetSmallIcon(Resource.Mipmap.appicon);
            builder.SetContentTitle(reminder.IsTodo ? "할일 알림" : "일정 알림");
            builder.SetContentText($"{reminder.Title} · {ReminderPlanner.Describe(reminder)}");
            builder.SetPriority((int)NotificationPriority.High);
            builder.SetAutoCancel(true);
            if (open is not null) builder.SetContentIntent(open);

            var notification = builder.Build();
            if (notification is not null) NotificationManagerCompat.From(context).Notify(reminder.AlarmId, notification);
        }
    }
}

/// <summary>Wakes on an alarm, says what is due, and books the next round.</summary>
[BroadcastReceiver(Enabled = true, Exported = false)]
public sealed class ReminderReceiver : BroadcastReceiver
{
    public override void OnReceive(Context? context, Intent? intent)
    {
        if (context is null) return;

        var pending = GoAsync();

        Task.Run(async () =>
        {
            try
            {
                await Workspace.EnsureLoadedAsync();
                ReminderScheduler.PostDue(context);
                ReminderScheduler.Reschedule(context);
                FlowdeckWidgets.RefreshAll(context);
            }
            catch (Exception)
            {
                // A receiver that throws takes the process with it, and there is nothing
                // useful to do about a reminder that could not be read.
            }
            finally
            {
                pending?.Finish();
            }
        });
    }
}

/// <summary>
/// Alarms do not survive a restart, so they are laid down again once the phone is back.
/// Without this, reminders would quietly stop working after every reboot.
/// </summary>
[BroadcastReceiver(Enabled = true, Exported = true)]
[IntentFilter(new[] { Intent.ActionBootCompleted, Intent.ActionMyPackageReplaced })]
public sealed class BootReceiver : BroadcastReceiver
{
    public override void OnReceive(Context? context, Intent? intent)
    {
        if (context is null) return;

        var pending = GoAsync();

        Task.Run(async () =>
        {
            try
            {
                await Workspace.EnsureLoadedAsync();
                ReminderScheduler.Reschedule(context);
                FlowdeckWidgets.RefreshAll(context);
            }
            catch (Exception)
            {
            }
            finally
            {
                pending?.Finish();
            }
        });
    }
}

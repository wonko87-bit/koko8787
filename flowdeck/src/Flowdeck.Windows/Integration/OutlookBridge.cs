using System.Runtime.InteropServices;
using System.Threading;
using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;

namespace Flowdeck.Windows.Integration;

/// <summary>
/// Writes entries into the copy of Outlook installed on this machine, over COM.
///
/// Chosen over the Graph API because it needs nothing granted: no app registration, no
/// admin consent, no network path. It runs as the signed-in user against data Outlook has
/// already synchronised, so on a locked-down corporate desktop it is usually the only
/// route that will be approved — and Outlook carries the entry up to the server itself.
///
/// Bound late, through <c>dynamic</c> and the ProgID, rather than against an Office
/// interop assembly. That keeps the build free of an Office dependency, works with
/// whichever Outlook version is present, and lets a machine without Outlook simply report
/// itself unavailable instead of failing to load.
/// </summary>
public sealed class OutlookBridge : IExternalStore
{
    // Outlook object model constants. Named here rather than referenced from an interop
    // assembly, which is the whole point of binding late.
    private const int OlAppointmentItem = 1;
    private const int OlTaskItem = 3;

    private const int OlRecursDaily = 0;
    private const int OlRecursWeekly = 1;
    private const int OlRecursMonthly = 2;
    private const int OlRecursYearly = 5;

    private const int OlImportanceLow = 0;
    private const int OlImportanceNormal = 1;
    private const int OlImportanceHigh = 2;

    public string Provider => "outlook";

    public string DisplayName => "Outlook";

    /// <summary>
    /// True when Outlook is registered on this machine. Deliberately only a registry
    /// lookup: starting Outlook to find out would be a rude thing to do on a check.
    /// </summary>
    public bool IsAvailable => Type.GetTypeFromProgID("Outlook.Application") is not null;

    public Task<ExternalLink> PushAsync(CalendarEvent calendarEvent, CancellationToken cancellationToken = default) =>
        RunOnStaAsync(() => CreateAppointment(calendarEvent), cancellationToken);

    public Task<ExternalLink> PushAsync(TodoItem todo, CancellationToken cancellationToken = default) =>
        RunOnStaAsync(() => CreateTask(todo), cancellationToken);

    // ---- appointments ------------------------------------------------------

    private ExternalLink CreateAppointment(CalendarEvent source)
    {
        dynamic? application = null;
        dynamic? item = null;

        try
        {
            application = CreateApplication();
            item = application.CreateItem(OlAppointmentItem);

            item.Subject = source.Title;
            if (!string.IsNullOrWhiteSpace(source.Location)) item.Location = source.Location;
            item.Body = BuildBody(source.Notes, source.SourceText);

            if (source.IsAllDay)
            {
                // Set the flag first: Outlook snaps the times to whole days as it is applied,
                // and its end is exclusive, unlike ours which runs to the last second.
                item.AllDayEvent = true;
                item.Start = source.Start.Date;
                item.End = source.End.Date.AddDays(1);
            }
            else
            {
                item.Start = source.Start;
                item.End = source.End > source.Start ? source.End : source.Start.AddHours(1);
            }

            if (source.Tags.Count > 0) item.Categories = string.Join(", ", source.Tags);

            if (source.ReminderMinutesBefore is > 0)
            {
                item.ReminderSet = true;
                item.ReminderMinutesBeforeStart = source.ReminderMinutesBefore.Value;
            }
            else
            {
                item.ReminderSet = false;
            }

            ApplyRecurrence(item, source.Recurrence, source.Start);

            item.Save();
            return LinkFrom(item);
        }
        finally
        {
            Release(item);
            Release(application);
        }
    }

    // ---- tasks -------------------------------------------------------------

    private ExternalLink CreateTask(TodoItem source)
    {
        dynamic? application = null;
        dynamic? item = null;

        try
        {
            application = CreateApplication();
            item = application.CreateItem(OlTaskItem);

            item.Subject = source.Title;
            item.Body = BuildBody(source.Notes, source.SourceText);

            if (source.DueAt.HasValue)
            {
                // Outlook tasks are due on a day, not at an hour; the time of day only
                // survives on the reminder.
                item.DueDate = source.DueAt.Value.Date;
                item.StartDate = source.DueAt.Value.Date;
            }

            item.Importance = source.Priority switch
            {
                Priority.Urgent or Priority.High => OlImportanceHigh,
                Priority.Low => OlImportanceLow,
                _ => OlImportanceNormal,
            };

            if (source.Tags.Count > 0) item.Categories = string.Join(", ", source.Tags);

            if (source.DueAt.HasValue && source.ReminderMinutesBefore is > 0)
            {
                item.ReminderSet = true;
                item.ReminderTime = source.DueAt.Value.AddMinutes(-source.ReminderMinutesBefore.Value);
            }

            if (source.DueAt.HasValue) ApplyRecurrence(item, source.Recurrence, source.DueAt.Value);

            item.Save();
            return LinkFrom(item);
        }
        finally
        {
            Release(item);
            Release(application);
        }
    }

    // ---- shared ------------------------------------------------------------

    private static dynamic CreateApplication()
    {
        var type = Type.GetTypeFromProgID("Outlook.Application")
                   ?? throw new InvalidOperationException("Outlook이 설치되어 있지 않습니다.");

        return Activator.CreateInstance(type)
               ?? throw new InvalidOperationException("Outlook을 시작하지 못했습니다.");
    }

    /// <summary>
    /// Maps our repeat rule onto Outlook's. Anything Outlook cannot express is left off
    /// rather than approximated — a wrong repeat is worse than none.
    /// </summary>
    private static void ApplyRecurrence(dynamic item, Recurrence recurrence, DateTime start)
    {
        if (!recurrence.IsRepeating) return;

        dynamic? pattern = null;
        try
        {
            pattern = item.GetRecurrencePattern();
            var interval = Math.Max(1, recurrence.Interval);

            switch (recurrence.Kind)
            {
                case RecurrenceKind.Daily:
                    pattern.RecurrenceType = OlRecursDaily;
                    pattern.Interval = interval;
                    break;

                case RecurrenceKind.Weekdays:
                    pattern.RecurrenceType = OlRecursWeekly;
                    pattern.Interval = 1;
                    pattern.DayOfWeekMask = DayMask(
                        new[]
                        {
                            DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
                            DayOfWeek.Thursday, DayOfWeek.Friday,
                        });
                    break;

                case RecurrenceKind.Weekly:
                    pattern.RecurrenceType = OlRecursWeekly;
                    pattern.Interval = interval;
                    pattern.DayOfWeekMask = DayMask(
                        recurrence.DaysOfWeek.Count > 0 ? recurrence.DaysOfWeek : new List<DayOfWeek> { start.DayOfWeek });
                    break;

                case RecurrenceKind.Monthly:
                    pattern.RecurrenceType = OlRecursMonthly;
                    pattern.Interval = interval;
                    pattern.DayOfMonth = start.Day;
                    break;

                case RecurrenceKind.Yearly:
                    pattern.RecurrenceType = OlRecursYearly;
                    // Outlook counts a yearly interval in months, and insists on multiples of 12.
                    pattern.Interval = 12 * interval;
                    pattern.MonthOfYear = start.Month;
                    pattern.DayOfMonth = start.Day;
                    break;

                default:
                    return;
            }

            pattern.PatternStartDate = start.Date;

            if (recurrence.Until.HasValue)
            {
                pattern.PatternEndDate = recurrence.Until.Value.Date;
            }
            else
            {
                pattern.NoEndDate = true;
            }
        }
        finally
        {
            Release(pattern);
        }
    }

    private static int DayMask(IEnumerable<DayOfWeek> days) =>
        days.Distinct().Sum(day => day switch
        {
            DayOfWeek.Sunday => 1,
            DayOfWeek.Monday => 2,
            DayOfWeek.Tuesday => 4,
            DayOfWeek.Wednesday => 8,
            DayOfWeek.Thursday => 16,
            DayOfWeek.Friday => 32,
            _ => 64,
        });

    private static string BuildBody(string notes, string sourceText)
    {
        var lines = new List<string>();
        if (!string.IsNullOrWhiteSpace(notes)) lines.Add(notes);
        if (!string.IsNullOrWhiteSpace(sourceText)) lines.Add($"[Flowdeck] {sourceText}");
        return string.Join(Environment.NewLine + Environment.NewLine, lines);
    }

    /// <summary>
    /// Records both identifiers Outlook needs to reopen the item later. The store id
    /// matters once more than one mailbox is connected, and a read-back mode will need it.
    /// </summary>
    private static ExternalLink LinkFrom(dynamic item)
    {
        string entryId = item.EntryID;
        string? storeId = null;

        dynamic? parent = null;
        try
        {
            parent = item.Parent;
            storeId = parent?.StoreID as string;
        }
        catch (Exception e) when (e is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or COMException)
        {
            // Not every folder exposes a StoreID; the EntryID alone is still useful.
        }
        finally
        {
            Release(parent);
        }

        return new ExternalLink
        {
            Provider = "outlook",
            EntryId = entryId,
            StoreId = storeId,
            SyncedAt = DateTime.Now,
        };
    }

    private static void Release(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject)) Marshal.ReleaseComObject(comObject);
    }

    /// <summary>
    /// Runs the work on a fresh single-threaded apartment.
    ///
    /// Outlook's object model is STA-bound, and attaching to it can take seconds if
    /// Outlook is not already running. Doing that on the UI thread would freeze the widget
    /// mid-capture, so each push gets its own short-lived apartment instead.
    /// </summary>
    private static Task<T> RunOnStaAsync<T>(Func<T> work, CancellationToken cancellationToken)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        if (cancellationToken.IsCancellationRequested)
        {
            completion.SetCanceled(cancellationToken);
            return completion.Task;
        }

        var thread = new Thread(() =>
        {
            try
            {
                completion.SetResult(work());
            }
            catch (Exception e)
            {
                completion.SetException(e);
            }
        })
        {
            IsBackground = true,
            Name = "Flowdeck.Outlook",
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return completion.Task;
    }
}

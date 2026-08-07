using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace Flowdeck.Windows.Services;

/// <summary>
/// Catches what would otherwise end the process without a word.
///
/// Flowdeck lives in the notification area, so an unhandled exception used to take the
/// whole app down with no window left behind to say what happened. Anything raised on the
/// UI thread is now logged, reported, and swallowed: a failed click is not a reason to
/// lose the session. Faults from the other two sources cannot be recovered from, but they
/// are at least written down.
/// </summary>
public static class CrashReporter
{
    private static string _logPath = string.Empty;
    private static Action<string, string>? _notify;

    public static string LogPath => _logPath;

    public static void Install(string dataFolder, Action<string, string>? notify = null)
    {
        _logPath = Path.Combine(dataFolder, "crash.log");
        _notify = notify;

        Application.Current.DispatcherUnhandledException += OnDispatcherException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    /// <summary>Records a fault the caller has already contained.</summary>
    public static void ReportRecovered(string context, Exception exception) =>
        Report(context, exception, recovered: true);

    /// <summary>Lets a fire-and-forget task report its fault instead of disappearing.</summary>
    public static void Observe(Task task, string context) =>
        task.ContinueWith(
            faulted => Report(context, faulted.Exception, recovered: true),
            CancellationToken.None,
            TaskContinuationOptions.OnlyOnFaulted,
            TaskScheduler.Default);

    private static void OnDispatcherException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        Report("UI", e.Exception, recovered: true);

        // Keep running. The alternative is losing the tray icon, the hot keys and the
        // widget because one command threw.
        e.Handled = true;
    }

    private static void OnDomainException(object sender, UnhandledExceptionEventArgs e) =>
        Report("AppDomain", e.ExceptionObject as Exception, recovered: false);

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        Report("Task", e.Exception, recovered: true);
        e.SetObserved();
    }

    private static void Report(string context, Exception? exception, bool recovered)
    {
        if (exception is null) return;

        Write(context, exception);

        if (!recovered) return;

        var summary = exception is AggregateException aggregate
            ? aggregate.Flatten().InnerExceptions.FirstOrDefault()?.Message ?? exception.Message
            : exception.Message;

        _notify?.Invoke("Flowdeck 오류", summary);
    }

    private static void Write(string context, Exception exception)
    {
        try
        {
            var directory = Path.GetDirectoryName(_logPath);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            var entry = new StringBuilder()
                .Append("=== ")
                .Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                .Append(" · ")
                .AppendLine(context)
                .AppendLine(exception.ToString())
                .AppendLine()
                .ToString();

            File.AppendAllText(_logPath, entry, Encoding.UTF8);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            // If even the log cannot be written there is nothing further to try.
        }
    }
}

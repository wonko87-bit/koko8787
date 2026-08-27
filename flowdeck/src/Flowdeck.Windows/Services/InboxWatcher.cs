using System.IO;
using System.Windows.Threading;
using Flowdeck.Core.Export;
using Flowdeck.Core.Services;

namespace Flowdeck.Windows.Services;

/// <summary>
/// Watches the inbox folder and hands anything that lands there to <see cref="InboxImporter"/>.
///
/// Only the watching lives here. The importer knows what a transfer file is and where a read
/// one goes afterwards; this class exists because a <c>FileSystemWatcher</c> calls back on a
/// thread-pool thread, and the repository it ends up talking to belongs to the UI thread.
/// </summary>
public sealed class InboxWatcher : IDisposable
{
    private readonly InboxImporter _importer;
    private readonly Dispatcher _dispatcher;
    private readonly Action<string, string>? _notify;

    /// <summary>
    /// Paths already on their way through. One arrival raises more than one event — a create
    /// and a write, or a rename — and the importer would otherwise queue up behind itself
    /// for each of them.
    /// </summary>
    private readonly HashSet<string> _inFlight = new(StringComparer.OrdinalIgnoreCase);

    private FileSystemWatcher? _watcher;
    private bool _disposed;

    public InboxWatcher(
        WorkspaceRepository repository,
        string folder,
        Dispatcher dispatcher,
        Action<string, string>? notify = null)
    {
        _dispatcher = dispatcher;
        _notify = notify;

        // Back onto the UI thread to import: the repository is not thread-safe and its
        // Changed event runs straight into the widget's bindings.
        _importer = new InboxImporter(
            folder,
            archive => _dispatcher.InvokeAsync(() => repository.ImportAsync(archive)).Task.Unwrap());
    }

    public string Folder => _importer.Folder;

    /// <summary>
    /// Creates the folders, takes whatever arrived while the app was closed, and then watches.
    /// The scan comes first so a file that arrives during startup is either found by it or
    /// reported by the watcher — and either way is skipped the second time by its id.
    /// </summary>
    public void Start()
    {
        _importer.EnsureFolders();

        // Off the UI thread: Start runs during startup, and reading whatever piled up while
        // the app was closed should not hold the first paint.
        CrashReporter.Observe(Task.Run(DrainAsync), "InboxWatcher.Drain");

        _watcher = new FileSystemWatcher(_importer.Folder, "*" + TransferFile.Extension)
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,

            // 처리됨 and 실패 live in here. Recursing would re-read everything already read.
            IncludeSubdirectories = false,
        };

        // The sending side is expected to build the file under a temporary name and rename
        // it into place, which is atomic and never exposes a half-written file. That arrives
        // as Renamed and not as Created, so a watcher listening only for Created sees nothing
        // at all. Created is here for a file copied in by hand.
        _watcher.Created += OnAppeared;
        _watcher.Renamed += OnAppeared;
        _watcher.Error += OnWatcherError;

        // Deleted is deliberately not handled: moving a file we have just read into 처리됨
        // looks like a deletion from in here, and reacting to it would be reacting to
        // ourselves.
        _watcher.EnableRaisingEvents = true;
    }

    private void OnAppeared(object sender, FileSystemEventArgs e) =>
        CrashReporter.Observe(ProcessAsync(e.FullPath), "InboxWatcher.Process");

    /// <summary>
    /// The watcher gave up — its buffer overflowed, or the folder went away. Rebuild it:
    /// without this the hand-over stops silently for the rest of the session.
    /// </summary>
    private void OnWatcherError(object sender, ErrorEventArgs e)
    {
        CrashReporter.ReportRecovered("InboxWatcher", e.GetException());

        _dispatcher.InvokeAsync(() =>
        {
            if (_disposed) return;

            _watcher?.Dispose();
            _watcher = null;

            try
            {
                Start();
            }
            catch (Exception restartFailed) when (restartFailed is IOException or ArgumentException)
            {
                // The folder is gone and could not be remade. Nothing further to try.
                CrashReporter.ReportRecovered("InboxWatcher.Restart", restartFailed);
            }
        });
    }

    private async Task DrainAsync()
    {
        foreach (var result in await _importer.DrainAsync())
        {
            Report(result);
        }
    }

    private async Task ProcessAsync(string path)
    {
        lock (_inFlight)
        {
            if (!_inFlight.Add(path)) return;
        }

        try
        {
            var result = await _importer.ProcessAsync(path);
            if (result is not null) Report(result);
        }
        finally
        {
            lock (_inFlight) _inFlight.Remove(path);
        }
    }

    /// <summary>
    /// Says what arrived. Nothing is said about a file that only contained entries already
    /// here, or one still being written — neither is something the user did or can act on.
    /// </summary>
    private void Report(InboxFileResult result)
    {
        switch (result.Outcome)
        {
            case InboxOutcome.Imported when result.Import is { Added: > 0 } import:
                Notify("Flowdeck", import.Describe());
                break;

            case InboxOutcome.Rejected:
                Notify("가져오지 못한 파일이 있습니다", $"{Path.GetFileName(result.Path)} · {result.Problem}");
                break;
        }
    }

    private void Notify(string title, string message) =>
        _dispatcher.InvokeAsync(() => _notify?.Invoke(title, message));

    public void Dispose()
    {
        _disposed = true;
        _watcher?.Dispose();
        _watcher = null;
    }
}

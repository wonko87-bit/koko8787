using Flowdeck.Core.Export;

namespace Flowdeck.Core.Services;

/// <summary>What happened to one file that was dropped in the inbox.</summary>
public enum InboxOutcome
{
    /// <summary>Read, imported and filed away. See <see cref="InboxFileResult.Import"/> for what it added.</summary>
    Imported,

    /// <summary>Not a transfer file, or too big to be one. Moved aside, never deleted.</summary>
    Rejected,

    /// <summary>Still locked by whoever is writing it. Left where it is for the next scan.</summary>
    Deferred,
}

/// <summary>One file's fate. <paramref name="Problem"/> is set for anything but a clean import.</summary>
public sealed record InboxFileResult(
    string Path,
    InboxOutcome Outcome,
    ImportResult? Import = null,
    string? Problem = null);

/// <summary>
/// Takes entries handed over by another application through a folder.
///
/// The bargain is deliberately small: the other side writes a transfer file into the inbox
/// and never touches anything else of Flowdeck's — above all not <c>workspace.json</c>, which
/// a running Flowdeck would overwrite from memory the next time it saved. Everything the
/// hand-over needs is already here, so nothing new was invented for it: the file is the same
/// format the export button writes, and <see cref="WorkspaceRepository.ImportAsync"/> does the
/// reading and the duplicate check.
///
/// Nothing is ever deleted. A file that was read moves to <see cref="DoneFolderName"/> and a
/// file that could not be read moves to <see cref="FailedFolderName"/>, both inside the inbox,
/// so a person can see what arrived and a watcher pointed at the inbox is not woken by the
/// move. Deleting would be tidier and cannot be undone.
///
/// The file system parts live here rather than in the desktop layer so they can be tested;
/// the <c>FileSystemWatcher</c> that calls in belongs to the app, which knows how to get back
/// onto its UI thread.
/// </summary>
public sealed class InboxImporter
{
    public const string DoneFolderName = "처리됨";

    public const string FailedFolderName = "실패";

    /// <summary>
    /// Above this, the file is not a hand-over that went wrong — it is something else that
    /// landed in the folder, and reading it into memory would be the only harm done.
    /// </summary>
    public const long SizeCeiling = 4 * 1024 * 1024;

    /// <summary>
    /// How many times to come back to a file that is locked. The writer is expected to build
    /// the file under another name and rename it into place, which is atomic and never shows
    /// a half-written file — this is for the copy-and-paste case, where it is not.
    /// </summary>
    private const int LockedAttempts = 5;

    private readonly Func<TransferArchive, Task<ImportResult>> _import;

    /// <summary>
    /// One file at a time, whichever way they arrive. Two imports at once would each save the
    /// workspace from its own copy in memory and the later save would drop the earlier one.
    /// </summary>
    private readonly SemaphoreSlim _gate = new(1, 1);

    public InboxImporter(string folder, Func<TransferArchive, Task<ImportResult>> import)
    {
        Folder = folder;
        _import = import;
        DoneFolder = Path.Combine(folder, DoneFolderName);
        FailedFolder = Path.Combine(folder, FailedFolderName);
    }

    /// <summary>Watched for arrivals. Written by the other application, never by Flowdeck.</summary>
    public string Folder { get; }

    public string DoneFolder { get; }

    public string FailedFolder { get; }

    /// <summary>Waits between attempts at a locked file. Replaced by tests, which will not wait.</summary>
    public Func<TimeSpan, Task> Delay { get; set; } = duration => Task.Delay(duration);

    public void EnsureFolders()
    {
        Directory.CreateDirectory(Folder);
        Directory.CreateDirectory(DoneFolder);
        Directory.CreateDirectory(FailedFolder);
    }

    /// <summary>
    /// Files waiting at the top of the inbox, oldest first. Called at startup for whatever
    /// arrived while Flowdeck was closed — the hand-over is not expected to be live, and a
    /// sender has no way of knowing whether anything is running.
    /// </summary>
    public IReadOnlyList<string> Pending()
    {
        if (!Directory.Exists(Folder)) return Array.Empty<string>();

        return Directory.EnumerateFiles(Folder, "*" + TransferFile.Extension, SearchOption.TopDirectoryOnly)
            .Where(IsTransferName)
            .OrderBy(File.GetLastWriteTimeUtc)
            .ToList();
    }

    /// <summary>Everything waiting, in order. Returns one result per file it got to.</summary>
    public async Task<IReadOnlyList<InboxFileResult>> DrainAsync()
    {
        var results = new List<InboxFileResult>();
        foreach (var path in Pending())
        {
            var result = await ProcessAsync(path).ConfigureAwait(false);
            if (result is not null) results.Add(result);
        }

        return results;
    }

    /// <summary>
    /// Reads one file and files it away. Returns null when there is nothing to report: the
    /// file is gone, or is not one of ours. A watcher raises several events for one arrival,
    /// so being called twice for the same file has to be ordinary rather than an error.
    /// </summary>
    public async Task<InboxFileResult?> ProcessAsync(string path)
    {
        // The pattern "*.txt" is matched against short names as well, which lets a
        // .txtx through; and a watcher hands over whatever changed, ours or not.
        if (!IsTransferName(path)) return null;

        await _gate.WaitAsync().ConfigureAwait(false);
        try
        {
            var file = new FileInfo(path);
            if (!file.Exists) return null;

            if (file.Length > SizeCeiling)
            {
                Move(path, FailedFolder);
                return new InboxFileResult(path, InboxOutcome.Rejected, Problem: "파일이 너무 큽니다.");
            }

            var contents = await ReadAsync(path).ConfigureAwait(false);
            if (contents is null)
            {
                // Whoever holds it is still writing. Leave it: the next startup scan
                // will find it, and moving a half-written file would lose it.
                return new InboxFileResult(path, InboxOutcome.Deferred, Problem: "파일이 잠겨 있어 다음에 다시 시도합니다.");
            }

            TransferArchive archive;
            try
            {
                archive = TransferFile.Read(contents);
            }
            catch (FormatException e)
            {
                Move(path, FailedFolder);
                return new InboxFileResult(path, InboxOutcome.Rejected, Problem: e.Message);
            }

            var imported = await _import(archive).ConfigureAwait(false);

            // Filed away only once the entries are saved: a crash between the two leaves the
            // file where it is, and importing it again is harmless — the ids already match.
            Move(path, DoneFolder);
            return new InboxFileResult(path, InboxOutcome.Imported, imported);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>The file's contents, or null if it stayed locked for every attempt.</summary>
    private async Task<string?> ReadAsync(string path)
    {
        var wait = TimeSpan.FromMilliseconds(200);

        for (var attempt = 0; ; attempt++)
        {
            try
            {
                return File.ReadAllText(path);
            }
            catch (IOException) when (attempt < LockedAttempts)
            {
                await Delay(wait).ConfigureAwait(false);
                wait += wait;
            }
            catch (IOException)
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Moves a file into one of the inbox's own sub-folders, never over something already
    /// there. Two hand-overs a second apart can carry the same name, and the older one is
    /// not worth less than the newer.
    /// </summary>
    private static void Move(string path, string destinationFolder)
    {
        try
        {
            Directory.CreateDirectory(destinationFolder);
            File.Move(path, Unused(destinationFolder, Path.GetFileName(path)));
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            // The entries are already in. Failing to tidy up afterwards is not worth
            // reporting, and the file will simply be seen again — and skipped — next time.
        }
    }

    private static string Unused(string folder, string fileName)
    {
        var candidate = Path.Combine(folder, fileName);
        if (!File.Exists(candidate)) return candidate;

        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);

        for (var n = 2; ; n++)
        {
            candidate = Path.Combine(folder, $"{stem} ({n}){extension}");
            if (!File.Exists(candidate)) return candidate;
        }
    }

    private static bool IsTransferName(string path) =>
        string.Equals(Path.GetExtension(path), TransferFile.Extension, StringComparison.OrdinalIgnoreCase);
}

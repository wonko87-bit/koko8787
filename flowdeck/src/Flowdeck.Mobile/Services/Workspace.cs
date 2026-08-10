using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Settings;
using Flowdeck.Core.Storage;

namespace Flowdeck.Mobile.Services;

/// <summary>
/// Everything the pages share: the workspace, the parser and the settings behind them.
///
/// A static holder rather than a container. There is exactly one of each for the life of
/// the process, the pages are created by Shell rather than by us, and loading has to happen
/// once before any page shows something — which a plain <see cref="EnsureLoadedAsync"/>
/// expresses more honestly than a registration graph would.
/// </summary>
public static class Workspace
{
    private static readonly SemaphoreSlim Gate = new(1, 1);
    private static bool _loaded;

    public static string DataFolder => FileSystem.AppDataDirectory;

    public static string SettingsPath => Path.Combine(DataFolder, "settings.json");

    public static AppSettings Settings { get; private set; } = new();

    public static WorkspaceRepository Repository { get; private set; } =
        new(new JsonWorkspaceStore(Path.Combine(FileSystem.AppDataDirectory, "workspace.json")));

    public static NaturalLanguageParser Parser { get; private set; } = new();

    /// <summary>
    /// Reads settings and workspace off disk, once. Pages call this from OnAppearing rather
    /// than the app blocking on it at startup: a first frame that arrives late looks like a
    /// slow app, and one that never arrives because the UI thread is waiting on file IO
    /// looks like a broken one.
    /// </summary>
    public static async Task EnsureLoadedAsync()
    {
        if (_loaded) return;

        await Gate.WaitAsync();
        try
        {
            if (_loaded) return;

            Directory.CreateDirectory(DataFolder);

            Settings = AppSettings.LoadFrom(SettingsPath);
            Parser = new NaturalLanguageParser(Settings.Routing);
            ApplyParserSettings();

            await Repository.LoadAsync();
            Repository.Changed += (_, _) => NotifyEntriesChanged();
            _loaded = true;
        }
        finally
        {
            Gate.Release();
        }
    }

    /// <summary>
    /// Pushes the settings that the parser reads. Called after anything on the settings page
    /// changes, so the very next line typed is read the new way.
    /// </summary>
    public static void ApplyParserSettings()
    {
        Parser.AssumeAfternoonForBareHours = Settings.AssumeAfternoonForBareHours;
        Parser.FirstDayOfWeek = Settings.FirstDayOfWeek;
        Parser.Contacts = Settings.Contacts;

        // There is no Outlook here to talk to over COM, so an entry must never be marked as
        // heading for one — it would only ever come back as a failure the user cannot act on.
        //
        // !TM is left working. Clearing the markers would leave "!TM" sitting in the title
        // of anything typed out of habit, and the link opens the Teams app on a phone just
        // as it opens it on a desktop, so removing it would cost a feature to save nothing.
        Parser.PushExternalByDefault = false;
    }

    /// <summary>
    /// Set by the Android layer at startup. Called after anything changes so alarms and
    /// home-screen widgets can be brought back into line — the shared code raises it
    /// without knowing what, if anything, is listening.
    /// </summary>
    public static Action? EntriesChanged { get; set; }

    /// <summary>
    /// Which tab to land on, left by whatever launched the app. Read and cleared once the
    /// shell exists — a widget tap arrives long before there is anything to navigate.
    /// </summary>
    public static string? PendingRoute { get; set; }

    public static void NotifyEntriesChanged()
    {
        try
        {
            EntriesChanged?.Invoke();
        }
        catch (Exception)
        {
            // Rescheduling an alarm is never worth taking the app down for.
        }
    }

    public static void SaveSettings()
    {
        try
        {
            Settings.SaveTo(SettingsPath);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            // A setting that failed to persist is worth far less than a crash on a phone.
        }
    }
}

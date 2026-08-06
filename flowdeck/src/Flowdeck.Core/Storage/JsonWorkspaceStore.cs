using System.Text.Json;
using System.Text.Json.Serialization;

namespace Flowdeck.Core.Storage;

/// <summary>
/// Stores the workspace as one indented JSON file.
///
/// Writes go to a sibling temp file and are then swapped in, so a crash mid-write
/// leaves the previous file intact rather than a half-written one. The previous
/// contents are kept alongside as <c>.bak</c>.
/// </summary>
public sealed class JsonWorkspaceStore : IWorkspaceStore
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    private readonly SemaphoreSlim _gate = new(1, 1);

    public JsonWorkspaceStore(string filePath) => FilePath = filePath;

    public string FilePath { get; }

    /// <summary>The default location: <c>%APPDATA%\Flowdeck\workspace.json</c> on Windows.</summary>
    public static string DefaultFilePath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Flowdeck",
            "workspace.json");

    public async Task<Workspace> LoadAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!File.Exists(FilePath)) return new Workspace();

            await using var stream = File.OpenRead(FilePath);
            var workspace = await JsonSerializer
                .DeserializeAsync<Workspace>(stream, Options, cancellationToken)
                .ConfigureAwait(false);
            return workspace ?? new Workspace();
        }
        catch (JsonException)
        {
            // A corrupt file should not take the app down with it. Step aside and start clean;
            // the unreadable file is kept so nothing is silently thrown away.
            TryPreserveCorruptFile();
            return new Workspace();
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task SaveAsync(Workspace workspace, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var directory = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            var tempPath = FilePath + ".tmp";
            await using (var stream = File.Create(tempPath))
            {
                await JsonSerializer
                    .SerializeAsync(stream, workspace, Options, cancellationToken)
                    .ConfigureAwait(false);
            }

            if (File.Exists(FilePath))
            {
                File.Replace(tempPath, FilePath, FilePath + ".bak", ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tempPath, FilePath);
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    private void TryPreserveCorruptFile()
    {
        try
        {
            if (File.Exists(FilePath)) File.Copy(FilePath, FilePath + ".corrupt", overwrite: true);
        }
        catch (IOException)
        {
            // Nothing useful to do here; loading an empty workspace is still the right outcome.
        }
    }
}

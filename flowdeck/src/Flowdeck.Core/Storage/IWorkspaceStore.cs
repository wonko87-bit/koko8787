namespace Flowdeck.Core.Storage;

/// <summary>
/// Persistence boundary. The desktop app uses <see cref="JsonWorkspaceStore"/>; a mobile
/// build can drop in a SQLite or cloud-backed implementation without touching anything else.
/// </summary>
public interface IWorkspaceStore
{
    Task<Workspace> LoadAsync(CancellationToken cancellationToken = default);

    Task SaveAsync(Workspace workspace, CancellationToken cancellationToken = default);
}

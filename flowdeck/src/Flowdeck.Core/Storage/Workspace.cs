using Flowdeck.Core.Models;

namespace Flowdeck.Core.Storage;

/// <summary>Everything the app persists: one file, one object.</summary>
public sealed class Workspace
{
    /// <summary>Bumped when the on-disk shape changes, so a future version can migrate.</summary>
    public int SchemaVersion { get; set; } = 1;

    public List<TodoItem> Todos { get; set; } = new();

    public List<CalendarEvent> Events { get; set; } = new();
}

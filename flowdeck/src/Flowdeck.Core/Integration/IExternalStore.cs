using Flowdeck.Core.Models;

namespace Flowdeck.Core.Integration;

/// <summary>
/// Somewhere entries are copied to besides the local workspace.
///
/// Only pushing is defined here, which is all the desktop app does today: Flowdeck writes
/// to Outlook and never reads back, so there is no reconciliation to get wrong. Reading is
/// a separate interface added when it is wanted, implemented by the same class — the
/// records already carry an <see cref="ExternalLink"/> for it to match on.
///
/// Deliberately free of any Outlook or Windows type, so the core stays buildable
/// everywhere and a Graph implementation can slot in beside the COM one.
/// </summary>
public interface IExternalStore
{
    /// <summary>Stable key written into <see cref="ExternalLink.Provider"/>.</summary>
    string Provider { get; }

    /// <summary>Shown in settings, e.g. "Outlook".</summary>
    string DisplayName { get; }

    /// <summary>
    /// False when the target is not installed or reachable. Checked before pushing so a
    /// machine without Outlook simply saves locally instead of raising errors.
    /// </summary>
    bool IsAvailable { get; }

    /// <summary>Creates the event and returns where it landed. Throws if it could not.</summary>
    Task<ExternalLink> PushAsync(CalendarEvent calendarEvent, CancellationToken cancellationToken = default);

    /// <summary>Creates the todo and returns where it landed. Throws if it could not.</summary>
    Task<ExternalLink> PushAsync(TodoItem todo, CancellationToken cancellationToken = default);
}

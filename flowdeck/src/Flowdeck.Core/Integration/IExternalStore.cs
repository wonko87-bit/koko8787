using Flowdeck.Core.Models;

namespace Flowdeck.Core.Integration;

/// <summary>What became of the copy outside Flowdeck while a change was being saved.</summary>
public enum ExternalSync
{
    /// <summary>Nothing to do: the entry has no copy, or the integration is off.</summary>
    None,

    /// <summary>The copy was brought back in line.</summary>
    Updated,

    /// <summary>The copy was removed.</summary>
    Removed,

    /// <summary>
    /// The copy could not be found. It was deleted in Outlook, or moved between folders —
    /// which changes the identifier Outlook hands out, so the two are indistinguishable
    /// from here. The link is dropped and the entry carries on as a local one.
    /// </summary>
    Lost,

    /// <summary>Outlook refused, or could not be reached at all.</summary>
    Failed,
}

/// <summary>
/// Somewhere entries are copied to besides the local workspace.
///
/// Writing only, in one direction: Flowdeck is the master and never reads Outlook back, so
/// there is no reconciliation to get wrong. What it does do is keep its own copies in step
/// — an entry edited or deleted here is edited or deleted there — which the
/// <see cref="ExternalLink"/> on each record makes possible.
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

    /// <summary>
    /// Rewrites the copy the event's link points at. False when there is nothing there any
    /// more; throws only when the store itself could not be reached.
    /// </summary>
    Task<bool> UpdateAsync(CalendarEvent calendarEvent, CancellationToken cancellationToken = default);

    /// <summary>Rewrites the copy the todo's link points at. False when it is gone.</summary>
    Task<bool> UpdateAsync(TodoItem todo, CancellationToken cancellationToken = default);

    /// <summary>Removes the copy. False when it was already gone, which is not a failure.</summary>
    Task<bool> DeleteAsync(ExternalLink link, CancellationToken cancellationToken = default);
}

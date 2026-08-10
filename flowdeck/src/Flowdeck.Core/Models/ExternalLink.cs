namespace Flowdeck.Core.Models;

/// <summary>
/// Where a record ended up outside Flowdeck.
///
/// Push-only mode does not strictly need this, but reading back later does: without an
/// identity to match on, a second run would have no way to tell an entry it already sent
/// from a new one, and every read would duplicate the calendar. Stamping it now costs a
/// field and saves the mirror mode from being a rewrite.
/// </summary>
public sealed class ExternalLink
{
    /// <summary>Which system holds it — "outlook" today, room for "graph" later.</summary>
    public string Provider { get; set; } = string.Empty;

    /// <summary>The provider's own identifier. Outlook calls this the EntryID.</summary>
    public string EntryId { get; set; } = string.Empty;

    /// <summary>
    /// The store the entry lives in. Outlook needs this alongside the EntryID to reopen an
    /// item reliably when more than one mailbox is connected.
    /// </summary>
    public string? StoreId { get; set; }

    public DateTime SyncedAt { get; set; }
}

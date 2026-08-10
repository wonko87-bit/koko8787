namespace Flowdeck.Core.Parsing;

/// <summary>
/// Turns the names a person actually types into the addresses a meeting invite needs.
///
/// Deliberately a hand-kept list rather than a directory lookup: reading the company
/// address book would mean Graph, an app registration and someone's approval, which is
/// the whole thing this route exists to avoid. A dozen lines in settings.json covers the
/// handful of people anyone actually invites.
/// </summary>
public sealed class ContactBook
{
    /// <summary>Name as typed after the <c>@</c>, mapped to an email address.</summary>
    public Dictionary<string, string> Aliases { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Resolves one <c>@handle</c>. An address is already an address and passes through;
    /// anything else has to be in the book, and a name that is not returns null so the
    /// parser can put the text back where it found it.
    /// </summary>
    public string? Resolve(string handle)
    {
        if (string.IsNullOrWhiteSpace(handle)) return null;
        if (handle.Contains('@')) return handle;

        return Aliases.TryGetValue(handle, out var address) && !string.IsNullOrWhiteSpace(address)
            ? address
            : null;
    }
}

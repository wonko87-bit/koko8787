using System.Text.RegularExpressions;

namespace Flowdeck.Core.Parsing;

/// <summary>
/// Turns the names a person actually types into the addresses a meeting invite needs.
///
/// Deliberately a hand-kept list rather than a directory lookup: reading the company
/// address book would mean Graph, an app registration and someone's approval, which is
/// the whole thing this route exists to avoid. A dozen entries covers the handful of
/// people anyone actually invites.
/// </summary>
public sealed class ContactBook
{
    /// <summary>
    /// What may follow the <c>@</c> in a line. The scanner builds its regex from this and
    /// the editor validates against it, so a name that is accepted is a name that resolves —
    /// there is no second definition to drift out of step with this one.
    /// </summary>
    public const string NamePattern = @"[가-힣A-Za-z][가-힣A-Za-z0-9._-]*";

    /// <summary>What counts as an email address, on the same terms.</summary>
    public const string AddressPattern = @"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}";

    private static readonly Regex NameShape =
        new(@"\A" + NamePattern + @"\z", RegexOptions.CultureInvariant);

    private static readonly Regex AddressShape =
        new(@"\A" + AddressPattern + @"\z", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    /// <summary>Name as typed after the <c>@</c>, mapped to an email address.</summary>
    public Dictionary<string, string> Aliases { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// True when <c>@name</c> would actually be picked up. Notably false for anything
    /// containing a space: "@김 대리" stops at the space, so the entry would sit in the
    /// book being silently never found.
    /// </summary>
    public static bool IsUsableName(string? name) =>
        !string.IsNullOrEmpty(name) && NameShape.IsMatch(name);

    public static bool IsAddress(string? value) =>
        !string.IsNullOrEmpty(value) && AddressShape.IsMatch(value);

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

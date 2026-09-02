using Flowdeck.Core.Parsing;

namespace Flowdeck.Core.Services;

/// <summary>
/// Finds the file an entry came from.
///
/// When another application hands work over through the inbox it has no way to point at the
/// file it was about: the import clears <c>ExternalLink</c>, which names an item in one
/// machine's Outlook and would mean nothing here. What survives is the note, so the sending
/// side writes the path into it as a line of its own:
///
/// <code>파일: C:\Users\andrew\Documents\관리함\보고서.pdf</code>
///
/// That line is written for a person to read first — it is worth something even if nothing
/// ever parses it — and this turns it back into a path so the entry can offer to open it.
/// A note nobody wrote that line into simply has no file, which is the ordinary case.
/// </summary>
public static class AttachedFile
{
    /// <summary>
    /// Extensions Flowdeck will not hand to the shell.
    ///
    /// The path arrives inside a file another program wrote, and running whatever it names
    /// is a larger power than "open the report this todo is about" needs. The inbox sits in
    /// the user's own data folder and nobody else can write there, so this is not a hole
    /// being closed — it is a door not being opened. Anything on this list can still be
    /// shown in its folder, where opening it is Explorer's business and the user's choice.
    /// </summary>
    private static readonly HashSet<string> NotOpenable = new(StringComparer.OrdinalIgnoreCase)
    {
        ".exe", ".com", ".scr", ".pif", ".bat", ".cmd", ".msi", ".msp", ".msc",
        ".ps1", ".psm1", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh", ".hta",
        ".cpl", ".reg", ".lnk", ".url", ".jar", ".dll",
    };

    /// <summary>
    /// The path named in the note, or null when it names none. The first such line wins, so
    /// a sender that puts a header above it still works. Read through the same outline the
    /// weekly report uses, so what counts as a <c>파일:</c> line is decided in one place.
    /// </summary>
    public static string? PathIn(string? notes) => NoteOutline.Parse(notes).Files.FirstOrDefault();

    /// <summary>Whether opening this would be running it rather than reading it.</summary>
    public static bool IsRunnable(string? path) =>
        !string.IsNullOrEmpty(path) && NotOpenable.Contains(Path.GetExtension(path));
}

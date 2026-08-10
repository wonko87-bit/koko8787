using System.Globalization;
using System.Text;
using Flowdeck.Core.Parsing;

namespace Flowdeck.Core.Integration;

/// <summary>
/// Builds the URL that opens the Teams scheduling form with an entry already typed into it.
///
/// A link, not an API call, and that is the point. Creating the meeting outright would mean
/// Graph — an app registration, delegated permissions, an administrator saying yes — and it
/// would send an invite the moment a line was misread. This opens the form filled in and
/// stops. Teams mints the meeting when the user presses its own save button, so a parse the
/// user disagrees with costs a glance, not an apology to the whole invite list.
/// </summary>
public static class TeamsDeepLink
{
    private const string NewMeeting = "https://teams.microsoft.com/l/meeting/new";

    /// <summary>
    /// The link for one parsed line, or null when the line asked for no meeting.
    ///
    /// Times are passed only when the line actually named one. An all-day entry has no
    /// hour to offer, and guessing one would put the user in front of a form that looks
    /// filled in but is wrong — worse than an empty field they can see is empty.
    /// </summary>
    public static string? NewMeetingFor(ParsedEntry entry)
    {
        if (entry is null || !entry.OpenTeamsMeeting || entry.IsEmpty) return null;

        var query = new List<(string Key, string Value)>();

        var subject = entry.Title == ParsedEntry.UntitledPlaceholder ? string.Empty : entry.Title;
        if (subject.Length > 0) query.Add(("subject", subject));

        if (entry.HasTime && entry.Start.HasValue)
        {
            query.Add(("startTime", Iso(entry.Start.Value)));
            query.Add(("endTime", Iso(entry.End ?? entry.Start.Value.AddHours(1))));
        }

        if (entry.Attendees.Count > 0) query.Add(("attendees", string.Join(",", entry.Attendees)));

        // Carries the original line through, so the form shows what was typed next to what
        // it was read as. Nobody has to remember which of the two they meant.
        if (entry.RawInput.Length > 0) query.Add(("content", "[Flowdeck] " + entry.RawInput));

        return query.Count == 0 ? NewMeeting : NewMeeting + "?" + Encode(query);
    }

    /// <summary>
    /// ISO 8601 with the machine's UTC offset. Teams reads the offset, so a meeting typed
    /// at three o'clock stays at three o'clock for whoever opens the form.
    /// </summary>
    private static string Iso(DateTime when) =>
        when.ToString("yyyy-MM-ddTHH:mm:sszzz", CultureInfo.InvariantCulture);

    private static string Encode(List<(string Key, string Value)> query)
    {
        var builder = new StringBuilder();

        foreach (var (key, value) in query)
        {
            if (builder.Length > 0) builder.Append('&');
            builder.Append(key).Append('=').Append(Uri.EscapeDataString(value));
        }

        return builder.ToString();
    }
}

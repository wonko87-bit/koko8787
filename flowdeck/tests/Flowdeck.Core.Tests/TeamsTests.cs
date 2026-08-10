using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

public class TeamsMarkerTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static ParsedEntry Parse(string input, NaturalLanguageParser? parser = null) =>
        (parser ?? new NaturalLanguageParser()).Parse(input, Now);

    [Theory]
    [InlineData("!TM 내일 오후 3시 기획 리뷰")]
    [InlineData("!tm 내일 오후 3시 기획 리뷰")]
    [InlineData("!팀즈 내일 오후 3시 기획 리뷰")]
    [InlineData("내일 오후 3시 기획 리뷰 !TM")]
    public void TheMarkerSetsTheFlagAndLeavesTheTitle(string input)
    {
        var entry = Parse(input);

        Assert.True(entry.OpenTeamsMeeting);
        Assert.Equal("기획 리뷰", entry.Title);
    }

    [Fact]
    public void WithoutTheMarkerNoFormOpens() =>
        Assert.False(Parse("내일 오후 3시 기획 리뷰").OpenTeamsMeeting);

    /// <summary>Three axes now, and the order they are typed in still does not matter.</summary>
    [Theory]
    [InlineData("!CD !TM 내일 오후 3시 기획 리뷰")]
    [InlineData("!TM !CD 내일 오후 3시 기획 리뷰")]
    public void TheMarkersCombineWithRouting(string input)
    {
        var entry = Parse(input);

        Assert.Equal(EntryTarget.Calendar, entry.Target);
        Assert.True(entry.OpenTeamsMeeting);
        Assert.DoesNotContain("!", entry.Title);
    }

    /// <summary>
    /// Teams puts the meeting on the Outlook calendar itself, so the standing "send
    /// everything" setting has to stand down or the slot is booked twice.
    /// </summary>
    [Fact]
    public void TeamsSuppressesTheAutomaticOutlookCopy()
    {
        var parser = new NaturalLanguageParser { PushExternalByDefault = true };

        Assert.False(Parse("!TM 내일 오후 3시 기획 리뷰", parser).PushExternal);
        Assert.True(Parse("내일 오후 3시 기획 리뷰", parser).PushExternal);
    }

    /// <summary>...but a user who typed both markers meant both.</summary>
    [Fact]
    public void AnExplicitOutlookMarkerStillWins() =>
        Assert.True(Parse("!TM !OL 내일 오후 3시 기획 리뷰").PushExternal);

    [Fact]
    public void AnOrdinaryWordStartingLikeTheMarkerIsLeftAlone()
    {
        var entry = Parse("!TMP 파일 정리");

        Assert.False(entry.OpenTeamsMeeting);
        Assert.Contains("TMP", entry.Title);
    }
}

public class AttendeeTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static NaturalLanguageParser WithContacts()
    {
        var parser = new NaturalLanguageParser();
        parser.Contacts.Aliases["홍길동"] = "hong@corp.com";
        parser.Contacts.Aliases["kim"] = "kim@corp.com";
        return parser;
    }

    private static ParsedEntry Parse(string input, NaturalLanguageParser? parser = null) =>
        (parser ?? new NaturalLanguageParser()).Parse(input, Now);

    [Theory]
    [InlineData("!TM 내일 3시 회의 @hong@corp.com")]
    [InlineData("!TM 내일 3시 회의 hong@corp.com")]
    public void AnAddressIsTakenAsWritten(string input)
    {
        var entry = Parse(input);

        Assert.Equal(new[] { "hong@corp.com" }, entry.Attendees);
        Assert.Equal("회의", entry.Title);
    }

    [Fact]
    public void ANameIsLookedUp()
    {
        var entry = Parse("!TM 내일 3시 기획 리뷰 @홍길동 @kim", WithContacts());

        Assert.Equal(new[] { "hong@corp.com", "kim@corp.com" }, entry.Attendees);
        Assert.Equal("기획 리뷰", entry.Title);
    }

    /// <summary>
    /// An unknown name is far more likely to be prose than a colleague, so it goes back
    /// into the title rather than vanishing out of it.
    /// </summary>
    [Fact]
    public void AnUnknownNameStaysInTheTitle()
    {
        var entry = Parse("!TM 내일 3시 @어딘가 앞에서 만나기", WithContacts());

        Assert.Empty(entry.Attendees);
        Assert.Contains("@어딘가", entry.Title);
    }

    /// <summary>
    /// Nothing but a meeting has anyone to invite, so an address in an ordinary line is
    /// left where it was written rather than quietly disappearing out of the title.
    /// </summary>
    [Theory]
    [InlineData("hong@corp.com 에게 회신", "hong@corp.com")]
    [InlineData("내일 3시 @홍길동 자리에서 논의", "@홍길동")]
    public void WithoutAMeetingAnAddressIsJustText(string input, string expected)
    {
        var entry = Parse(input, WithContacts());

        Assert.Empty(entry.Attendees);
        Assert.Contains(expected, entry.Title);
    }

    /// <summary>
    /// Ctrl+M has to re-read the line rather than set a flag: turning the meeting on is
    /// what turns "@홍길동" from a word in the title into someone to invite.
    /// </summary>
    [Fact]
    public void FlippingTheFlagByKeyboardPullsTheAttendeesOut()
    {
        var parser = WithContacts();
        const string line = "내일 3시 기획 리뷰 @홍길동";

        var asTyped = parser.Parse(line, Now);
        var asMeeting = parser.Parse(line, Now, forceTeams: true);

        Assert.Empty(asTyped.Attendees);
        Assert.Contains("@홍길동", asTyped.Title);

        Assert.Equal(new[] { "hong@corp.com" }, asMeeting.Attendees);
        Assert.Equal("기획 리뷰", asMeeting.Title);
    }

    [Fact]
    public void TheSameAddressTwiceIsInvitedOnce() =>
        Assert.Single(Parse("!TM 회의 @hong@corp.com @HONG@corp.com").Attendees);

    [Fact]
    public void AnEmailIsNotMistakenForAHandle()
    {
        var entry = Parse("!TM 회의 @hong@corp.com", WithContacts());

        Assert.Equal(new[] { "hong@corp.com" }, entry.Attendees);
        Assert.DoesNotContain("@", entry.Title);
    }
}

public class ContactBookTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    [Theory]
    [InlineData("홍길동")]
    [InlineData("kim")]
    [InlineData("kim.lee")]
    [InlineData("박PM")]
    public void UsableNames(string name) => Assert.True(ContactBook.IsUsableName(name));

    /// <summary>
    /// A name with a space is the trap: it saves happily and then never matches, because
    /// "@김 대리" stops at the space. The editor has to refuse it up front.
    /// </summary>
    [Theory]
    [InlineData("김 대리")]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("@홍길동")]
    [InlineData("1번")]
    public void UnusableNames(string name) => Assert.False(ContactBook.IsUsableName(name));

    [Theory]
    [InlineData("hong@corp.com", true)]
    [InlineData("a.b+c@sub.corp.co.kr", true)]
    [InlineData("hong@corp", false)]
    [InlineData("hong", false)]
    [InlineData("", false)]
    public void AddressShapes(string value, bool expected) =>
        Assert.Equal(expected, ContactBook.IsAddress(value));

    /// <summary>
    /// The editor and the scanner have to agree, or the editor cheerfully accepts entries
    /// that are never found. Both read the same pattern, and this is what holds them to it.
    /// </summary>
    [Theory]
    [InlineData("홍길동")]
    [InlineData("kim")]
    [InlineData("kim.lee")]
    [InlineData("박PM")]
    public void EveryNameTheEditorAcceptsIsANameTheParserFinds(string name)
    {
        var parser = new NaturalLanguageParser();
        parser.Contacts.Aliases[name] = "someone@corp.com";

        var entry = parser.Parse($"!TM 내일 3시 회의 @{name}", Now);

        Assert.Equal(new[] { "someone@corp.com" }, entry.Attendees);
    }
}

public class TeamsDeepLinkTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static string Link(string input)
    {
        var parser = new NaturalLanguageParser();
        parser.Contacts.Aliases["홍길동"] = "hong@corp.com";
        return TeamsDeepLink.NewMeetingFor(parser.Parse(input, Now))!;
    }

    [Fact]
    public void NoMarkerMeansNoLink() =>
        Assert.Null(TeamsDeepLink.NewMeetingFor(new NaturalLanguageParser().Parse("내일 3시 회의", Now)));

    [Fact]
    public void TheLinkCarriesSubjectTimesAndAttendees()
    {
        var link = Link("!TM 8월 12일 오후 3시부터 4시까지 기획 리뷰 @홍길동");

        Assert.StartsWith("https://teams.microsoft.com/l/meeting/new?", link);
        Assert.Contains("subject=" + Uri.EscapeDataString("기획 리뷰"), link);
        Assert.Contains("startTime=" + Uri.EscapeDataString("2026-08-12T15:00:00"), link);
        Assert.Contains("endTime=" + Uri.EscapeDataString("2026-08-12T16:00:00"), link);
        Assert.Contains("attendees=" + Uri.EscapeDataString("hong@corp.com"), link);
    }

    [Fact]
    public void SeveralAttendeesAreCommaSeparated()
    {
        var link = Link("!TM 내일 3시 회의 a@corp.com b@corp.com");

        Assert.Contains("attendees=" + Uri.EscapeDataString("a@corp.com,b@corp.com"), link);
    }

    /// <summary>
    /// An all-day line has no hour to offer. Guessing one would hand the user a form that
    /// looks filled in and is wrong, which is worse than a field they can see is empty.
    /// </summary>
    [Fact]
    public void AnAllDayEntryLeavesTheTimeToTheUser()
    {
        var link = Link("!TM 8월 12일 기획 리뷰");

        Assert.Contains("subject=", link);
        Assert.DoesNotContain("startTime=", link);
        Assert.DoesNotContain("endTime=", link);
    }

    [Fact]
    public void TheOriginalLineTravelsInTheBody() =>
        Assert.Contains("content=" + Uri.EscapeDataString("[Flowdeck] !TM 내일 3시 회의"), Link("!TM 내일 3시 회의"));

    /// <summary>A subject with an ampersand must not split the query string.</summary>
    [Fact]
    public void AwkwardCharactersAreEscaped()
    {
        var link = Link("!TM 내일 3시 A&B 조율 #협의");

        Assert.DoesNotContain("A&B", link);
        Assert.Contains(Uri.EscapeDataString("A&B"), link);
    }

    /// <summary>
    /// Parses as one absolute URL — the check that matters, since this string is handed
    /// straight to the shell. Not IsWellFormedUriString, which also objects to characters
    /// that are escaped without strictly needing to be, such as the "!" of "!TM" in the body.
    /// </summary>
    [Fact]
    public void TheLinkIsOneAbsoluteUri()
    {
        Assert.True(Uri.TryCreate(Link("!TM 내일 3시 기획 리뷰 @홍길동"), UriKind.Absolute, out var uri));
        Assert.Equal(Uri.UriSchemeHttps, uri!.Scheme);
        Assert.Equal("teams.microsoft.com", uri.Host);
    }
}

public class TeamsCaptureTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    /// <summary>
    /// The entry is filed first and the form is only a link handed back afterwards, so a
    /// meeting the user then abandons in Teams still leaves the line in Flowdeck.
    /// </summary>
    [Fact]
    public async Task TheEntryIsSavedAndTheLinkComesBackWithIt()
    {
        var repo = new WorkspaceRepository(new InMemoryStore());

        var result = await repo.CaptureAsync(Parse("!CD !TM 내일 오후 3시 기획 리뷰"), Now);

        Assert.Single(repo.Events);
        Assert.NotNull(result.LaunchUrl);
        Assert.Contains("teams.microsoft.com", result.LaunchUrl);
    }

    [Fact]
    public async Task AnOrdinaryEntryHasNothingToOpen()
    {
        var repo = new WorkspaceRepository(new InMemoryStore());

        var result = await repo.CaptureAsync(Parse("내일 오후 3시 기획 리뷰"), Now);

        Assert.Null(result.LaunchUrl);
    }
}

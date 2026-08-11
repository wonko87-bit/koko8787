using Flowdeck.Core.Integration;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>Hands back whatever it was loaded with, and can be told to fail.</summary>
internal sealed class FakeCalendarReader : IExternalCalendarReader
{
    public string DisplayName => "Fake";

    public List<ExternalOccurrence> Occurrences { get; } = new();

    public bool ShouldFail { get; set; }

    public int Reads { get; private set; }

    public Task<IReadOnlyList<ExternalOccurrence>> ReadCalendarAsync(
        DateTime from,
        DateTime to,
        CancellationToken cancellationToken = default)
    {
        Reads++;
        if (ShouldFail) throw new InvalidOperationException("Outlook이 응답하지 않습니다");

        IReadOnlyList<ExternalOccurrence> inWindow = Occurrences
            .Where(o => o.End >= from && o.Start <= to)
            .ToList();

        return Task.FromResult(inWindow);
    }
}

/// <summary>
/// The work calendar drawn under the user's own entries. Read-only throughout: nothing here
/// may ever put something back, which is what keeps it free of who-wins questions.
/// </summary>
public class ExternalCalendarFeedTests
{
    private static readonly DateTime Now = new(2026, 8, 10, 9, 0, 0);
    private static readonly IReadOnlySet<string> Nothing = new HashSet<string>();

    private static ExternalOccurrence Meeting(
        string id,
        DateTime start,
        double hours = 1,
        bool allDay = false,
        string title = "회의") =>
        new(id, title, start, start.AddHours(hours), allDay, string.Empty);

    private static async Task<(ExternalCalendarFeed Feed, FakeCalendarReader Reader)> Load(
        params ExternalOccurrence[] occurrences)
    {
        var reader = new FakeCalendarReader();
        reader.Occurrences.AddRange(occurrences);

        var feed = new ExternalCalendarFeed(reader, () => Now) { IsEnabled = true };
        await feed.RefreshAsync(Now.Date.AddDays(-7), Now.Date.AddDays(30));
        return (feed, reader);
    }

    [Fact]
    public async Task ADayShowsWhatIsOnIt()
    {
        var (feed, _) = await Load(
            Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)),
            Meeting("b", new DateTime(2026, 8, 12, 14, 0, 0)));

        Assert.Single(feed.On(new DateTime(2026, 8, 11), Nothing));
        Assert.Empty(feed.On(new DateTime(2026, 8, 13), Nothing));
    }

    /// <summary>All-day first, then in start order, so a day reads top to bottom.</summary>
    [Fact]
    public async Task ADayIsOrderedForReading()
    {
        var day = new DateTime(2026, 8, 11);
        var (feed, _) = await Load(
            Meeting("late", day.AddHours(16), title: "오후"),
            Meeting("early", day.AddHours(9), title: "오전"),
            Meeting("all", day, hours: 24, allDay: true, title: "휴가"));

        Assert.Equal(
            new[] { "휴가", "오전", "오후" },
            feed.On(day, Nothing).Select(o => o.Title));
    }

    /// <summary>A meeting spanning days belongs to each of them.</summary>
    [Fact]
    public async Task AMultiDayMeetingShowsOnEveryDayItTouches()
    {
        var (feed, _) = await Load(
            Meeting("trip", new DateTime(2026, 8, 11, 9, 0, 0), hours: 48, title: "출장"));

        Assert.Single(feed.On(new DateTime(2026, 8, 11), Nothing));
        Assert.Single(feed.On(new DateTime(2026, 8, 12), Nothing));
        Assert.Single(feed.On(new DateTime(2026, 8, 13), Nothing));
        Assert.Empty(feed.On(new DateTime(2026, 8, 14), Nothing));
    }

    /// <summary>
    /// The one that matters. Entries sent with !OL come straight back out of Outlook, and
    /// without this they would be drawn twice — once as the user's row, once as a meeting.
    /// </summary>
    [Fact]
    public async Task WhatFlowdeckSentIsNotDrawnAgain()
    {
        var day = new DateTime(2026, 8, 11);
        var (feed, _) = await Load(
            Meeting("mine", day.AddHours(10), title: "내가 넣은 회의"),
            Meeting("theirs", day.AddHours(14), title: "남이 잡은 회의"));

        var mine = new HashSet<string> { "mine" };

        Assert.Equal(new[] { "남이 잡은 회의" }, feed.On(day, mine).Select(o => o.Title));
        Assert.DoesNotContain(day, feed.DaysWith(new HashSet<string> { "mine", "theirs" }));
    }

    /// <summary>The month dots have to count these, or a booked day reads as free.</summary>
    [Fact]
    public async Task TheDaysWithMeetingsAreReported()
    {
        var (feed, _) = await Load(
            Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)),
            Meeting("b", new DateTime(2026, 8, 20, 10, 0, 0)));

        Assert.Equal(
            new[] { new DateTime(2026, 8, 11), new DateTime(2026, 8, 20) },
            feed.DaysWith(Nothing).OrderBy(d => d));
    }

    /// <summary>Switched off, it shows nothing and reads nothing.</summary>
    [Fact]
    public async Task TurnedOffItIsSilent()
    {
        var reader = new FakeCalendarReader();
        reader.Occurrences.Add(Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)));

        var feed = new ExternalCalendarFeed(reader, () => Now) { IsEnabled = false };
        await feed.RefreshAsync(Now.Date, Now.Date.AddDays(30));

        Assert.Equal(0, reader.Reads);
        Assert.Empty(feed.On(new DateTime(2026, 8, 11), Nothing));
        Assert.Empty(feed.DaysWith(Nothing));
    }

    /// <summary>
    /// A calendar that empties itself every time Outlook is busy is worse than one that is a
    /// few minutes stale, so a failed read keeps the last good one and records why.
    /// </summary>
    [Fact]
    public async Task AFailedReadKeepsWhatWasAlreadyThere()
    {
        var (feed, reader) = await Load(Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)));
        reader.ShouldFail = true;

        await feed.RefreshAsync(Now.Date.AddDays(-7), Now.Date.AddDays(30));

        Assert.Single(feed.On(new DateTime(2026, 8, 11), Nothing));
        Assert.Equal("Outlook이 응답하지 않습니다", feed.LastError);
    }

    [Fact]
    public async Task AGoodReadClearsAnEarlierError()
    {
        var (feed, reader) = await Load(Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)));
        reader.ShouldFail = true;
        await feed.RefreshAsync(Now.Date, Now.Date.AddDays(30));

        reader.ShouldFail = false;
        await feed.RefreshAsync(Now.Date, Now.Date.AddDays(30));

        Assert.Null(feed.LastError);
    }

    [Fact]
    public async Task SwitchingItOffDropsWhatWasRead()
    {
        var (feed, _) = await Load(Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)));

        feed.Clear();

        Assert.Equal(0, feed.Count);
        Assert.Empty(feed.On(new DateTime(2026, 8, 11), Nothing));
    }

    /// <summary>Paging inside the window already read must not send us back to Outlook.</summary>
    [Fact]
    public async Task TheWindowIsKnownSoPagingDoesNotReread()
    {
        var (feed, _) = await Load(Meeting("a", new DateTime(2026, 8, 11, 10, 0, 0)));

        Assert.True(feed.Covers(Now.Date, Now.Date.AddDays(20)));
        Assert.False(feed.Covers(Now.Date, Now.Date.AddDays(60)));
    }

    /// <summary>Nothing to read from is not an error; the feature simply is not offered.</summary>
    [Fact]
    public async Task WithNoReaderItIsSimplyUnavailable()
    {
        var feed = new ExternalCalendarFeed(null, () => Now) { IsEnabled = true };

        await feed.RefreshAsync(Now.Date, Now.Date.AddDays(30));

        Assert.False(feed.IsAvailable);
        Assert.Empty(feed.On(Now.Date, Nothing));
        Assert.Null(feed.LastError);
    }

    /// <summary>
    /// End to end: an entry captured with !OL, read back out of the calendar, and correctly
    /// left out of the overlay because the repository knows it is one of ours.
    /// </summary>
    [Fact]
    public async Task AnEntryOfOursIsRecognisedThroughTheRepository()
    {
        var store = new FakeExternalStore();
        var repository = new WorkspaceRepository(new InMemoryStore(), store);
        await repository.CaptureAsync(
            new NaturalLanguageParser().Parse("!OL !CD 8월 11일 오전 10시 팀 회의", Now), Now);

        var entryId = repository.Events[0].ExternalLink!.EntryId;
        var (feed, _) = await Load(
            Meeting(entryId, new DateTime(2026, 8, 11, 10, 0, 0), title: "팀 회의"),
            Meeting("someone-else", new DateTime(2026, 8, 11, 15, 0, 0), title: "부서 회의"));

        var shown = feed.On(new DateTime(2026, 8, 11), repository.LinkedEntryIds);

        Assert.Equal(new[] { "부서 회의" }, shown.Select(o => o.Title));
    }
}

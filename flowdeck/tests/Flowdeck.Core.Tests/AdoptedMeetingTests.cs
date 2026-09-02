using Flowdeck.Core.Export;
using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// Taking a copy of somebody else's meeting so a note can be kept on it. The copy is the
/// user's; the meeting stays theirs, and nothing done to the one reaches the other.
/// </summary>
public class AdoptedMeetingTests
{
    private static readonly DateTime Now = new(2026, 8, 20, 9, 0, 0);
    private static readonly DateTime Monday = new(2026, 8, 24, 14, 0, 0);

    private static ExternalOccurrence Meeting(string id, DateTime start, string title = "설계 리뷰") =>
        new(id, title, start, start.AddHours(1), false, "3층 회의실");

    private static async Task<(WorkspaceRepository Repository, ExternalCalendarFeed Feed, FakeExternalStore Store)> Setup(
        params ExternalOccurrence[] occurrences)
    {
        var store = new FakeExternalStore();
        var repository = new WorkspaceRepository(new InMemoryStore(), store);
        await repository.LoadAsync();

        var reader = new FakeCalendarReader();
        reader.Occurrences.AddRange(occurrences);
        var feed = new ExternalCalendarFeed(reader, () => Now) { IsEnabled = true };
        await feed.RefreshAsync(Now.Date.AddDays(-7), Now.Date.AddDays(30));

        return (repository, feed, store);
    }

    // ---- the copy ------------------------------------------------------------

    [Fact]
    public async Task ACopyIsMadeAndTheOriginalLeavesTheOverlay()
    {
        var meeting = Meeting("M1", Monday);
        var (repository, feed, _) = await Setup(meeting);

        Assert.Single(feed.On(Monday, repository.Hides));

        var copy = await repository.AdoptAsync(meeting, Now);

        var stored = Assert.Single(repository.Events);
        Assert.Same(copy, stored);
        Assert.Equal("설계 리뷰", copy.Title);
        Assert.Equal("3층 회의실", copy.Location);
        Assert.Equal(Monday, copy.Start);
        Assert.Equal(Monday.AddHours(1), copy.End);
        Assert.Equal(new ExternalOrigin("M1", Monday), copy.Origin);
        Assert.Null(copy.ExternalLink);

        Assert.Empty(feed.On(Monday, repository.Hides));
        Assert.DoesNotContain(Monday.Date, feed.DaysWith(repository.Hides));
    }

    [Fact]
    public async Task AskingTwiceGivesTheSameCopy()
    {
        var meeting = Meeting("M1", Monday);
        var (repository, _, _) = await Setup(meeting);

        var first = await repository.AdoptAsync(meeting, Now);
        var second = await repository.AdoptAsync(meeting, Now.AddMinutes(5));

        Assert.Same(first, second);
        Assert.Single(repository.Events);
    }

    [Fact]
    public async Task TheOtherMorningsOfARepeatingMeetingStayVisible()
    {
        // Every occurrence of a repeat shares one entry id. Taking this Monday's must not
        // make next Monday's vanish.
        var thisWeek = Meeting("WEEKLY", Monday, "주간 회의");
        var nextWeek = Meeting("WEEKLY", Monday.AddDays(7), "주간 회의");
        var (repository, feed, _) = await Setup(thisWeek, nextWeek);

        await repository.AdoptAsync(thisWeek, Now);

        Assert.Empty(feed.On(Monday, repository.Hides));
        Assert.Single(feed.On(Monday.AddDays(7), repository.Hides));
        Assert.Contains(Monday.AddDays(7).Date, feed.DaysWith(repository.Hides));
    }

    [Fact]
    public async Task AMeetingWithNoTitleGetsOne()
    {
        var (repository, _, _) = await Setup();

        var copy = await repository.AdoptAsync(new ExternalOccurrence("M1", "  ", Monday, Monday.AddHours(1), false, ""), Now);

        Assert.Equal("(제목 없음)", copy.Title);
    }

    // ---- nothing goes back -----------------------------------------------------

    [Fact]
    public async Task ANoteWrittenOnTheCopyNeverReachesTheMeeting()
    {
        var meeting = Meeting("M1", Monday);
        var (repository, _, store) = await Setup(meeting);
        var copy = await repository.AdoptAsync(meeting, Now);

        var saved = await repository.UpdateEventAsync(copy.Id, new EntryEdit
        {
            Title = copy.Title,
            Notes = "결정: 코어 재질 B안으로 확정",
            When = copy.Start,
            HasTime = true,
        }, Now);

        Assert.True(saved.Found);
        Assert.Equal(ExternalSync.None, saved.External);
        Assert.Empty(store.UpdatedEvents);
        Assert.Equal("결정: 코어 재질 B안으로 확정", Assert.Single(repository.Events).Notes);
    }

    [Fact]
    public async Task DeletingTheCopyLeavesTheMeetingAloneAndBringsItBack()
    {
        var meeting = Meeting("M1", Monday);
        var (repository, feed, store) = await Setup(meeting);
        var copy = await repository.AdoptAsync(meeting, Now);

        await repository.DeleteEventAsync(copy.Id);

        Assert.Empty(store.Deleted);
        Assert.Empty(repository.Events);
        Assert.Single(feed.On(Monday, repository.Hides));
    }

    [Fact]
    public async Task MovingTheCopyDoesNotUnhideTheOriginal()
    {
        // The origin remembers the occurrence's own start, not the copy's current one.
        var meeting = Meeting("M1", Monday);
        var (repository, feed, _) = await Setup(meeting);
        var copy = await repository.AdoptAsync(meeting, Now);

        await repository.UpdateEventAsync(copy.Id, new EntryEdit
        {
            Title = copy.Title,
            When = Monday.AddHours(2),
            HasTime = true,
        }, Now);

        Assert.Empty(feed.On(Monday, repository.Hides));
    }

    // ---- it stays on this machine ------------------------------------------------

    [Fact]
    public async Task TheOriginDoesNotTravelInATransferFile()
    {
        var meeting = Meeting("M1", Monday);
        var (repository, _, _) = await Setup(meeting);
        var copy = await repository.AdoptAsync(meeting, Now);

        var text = TransferFile.Write(Array.Empty<TodoItem>(), new[] { copy }, DateTimeOffset.Now);
        var arrived = Assert.Single(TransferFile.Read(text).Events);

        Assert.Null(arrived.Origin);
        Assert.NotNull(copy.Origin);
    }

    [Fact]
    public async Task AnOriginArrivingFromElsewhereIsDropped()
    {
        var (repository, _, _) = await Setup();

        await repository.ImportAsync(new TransferArchive
        {
            Events = { new CalendarEvent { Title = "저쪽 회의", Start = Monday, End = Monday.AddHours(1), Origin = new ExternalOrigin("THEIRS", Monday) } },
        });

        Assert.Null(Assert.Single(repository.Events).Origin);
    }
}

using Flowdeck.Core.Integration;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>Records what it was asked to push, and can be told to fail.</summary>
internal sealed class FakeExternalStore : IExternalStore
{
    public string Provider => "fake";

    public string DisplayName => "Fake";

    public bool IsAvailable { get; set; } = true;

    public bool ShouldFail { get; set; }

    public List<string> PushedEvents { get; } = new();

    public List<string> PushedTodos { get; } = new();

    public Task<ExternalLink> PushAsync(CalendarEvent calendarEvent, CancellationToken cancellationToken = default)
    {
        if (ShouldFail) throw new InvalidOperationException("Outlook이 응답하지 않습니다");

        PushedEvents.Add(calendarEvent.Title);
        return Task.FromResult(new ExternalLink { Provider = "fake", EntryId = "e" + PushedEvents.Count });
    }

    public Task<ExternalLink> PushAsync(TodoItem todo, CancellationToken cancellationToken = default)
    {
        if (ShouldFail) throw new InvalidOperationException("Outlook이 응답하지 않습니다");

        PushedTodos.Add(todo.Title);
        return Task.FromResult(new ExternalLink { Provider = "fake", EntryId = "t" + PushedTodos.Count });
    }
}

public class ExternalMarkerTests
{
    private static readonly DateTime Now = new(2026, 8, 7, 10, 0, 0);

    private static ParsedEntry Parse(string input, bool byDefault = false) =>
        new NaturalLanguageParser { PushExternalByDefault = byDefault }.Parse(input, Now);

    [Theory]
    [InlineData("!OL 내일 오후 3시 회의")]
    [InlineData("!ol 내일 오후 3시 회의")]
    [InlineData("!아웃룩 내일 오후 3시 회의")]
    [InlineData("내일 오후 3시 회의 !OL")]
    public void OutlookMarkerSetsTheFlagAndLeavesTheTitle(string input)
    {
        var entry = Parse(input);

        Assert.True(entry.PushExternal);
        Assert.Equal("회의", entry.Title);
    }

    [Fact]
    public void WithoutAMarkerNothingIsPushed() =>
        Assert.False(Parse("내일 오후 3시 회의").PushExternal);

    /// <summary>
    /// The Outlook flag and the routing marker sit on different axes, so both survive
    /// together — and in either order, since each may take the head or the tail.
    /// </summary>
    [Theory]
    [InlineData("!CD !OL 내일 오후 3시 회의", EntryTarget.Calendar)]
    [InlineData("!OL !CD 내일 오후 3시 회의", EntryTarget.Calendar)]
    [InlineData("!TD !OL 금요일까지 보고서 제출", EntryTarget.Todo)]
    [InlineData("!OL !TD 금요일까지 보고서 제출", EntryTarget.Todo)]
    public void MarkersCombine(string input, EntryTarget expected)
    {
        var entry = Parse(input);

        Assert.Equal(expected, entry.Target);
        Assert.True(entry.PushExternal);
        Assert.DoesNotContain("!", entry.Title);
    }

    [Fact]
    public void TheDefaultCanBeTurnedOnForEveryEntry() =>
        Assert.True(Parse("내일 오후 3시 회의", byDefault: true).PushExternal);

    [Theory]
    [InlineData("!NOL 내일 오후 3시 회의")]
    [InlineData("내일 오후 3시 회의 !NOL")]
    [InlineData("!로컬 내일 오후 3시 회의")]
    public void LocalOnlyMarkerOverridesTheDefault(string input)
    {
        var entry = Parse(input, byDefault: true);

        Assert.False(entry.PushExternal);
        Assert.Equal("회의", entry.Title);
    }

    [Fact]
    public void AnOrdinaryWordStartingLikeAMarkerIsLeftAlone()
    {
        var entry = Parse("!OLD 자료 정리");

        Assert.False(entry.PushExternal);
        Assert.Contains("OLD", entry.Title);
    }
}

public class ExternalPushTests
{
    private static readonly DateTime Now = new(2026, 8, 7, 10, 0, 0);

    private static (WorkspaceRepository Repo, FakeExternalStore External) Build()
    {
        var external = new FakeExternalStore();
        return (new WorkspaceRepository(new InMemoryStore(), external), external);
    }

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    [Fact]
    public async Task BothTargetsPushAnEventAndATodo()
    {
        var (repo, external) = Build();

        await repo.CaptureAsync(Parse("!OL 내일 오후 9시 운동"), Now);

        Assert.Equal(new[] { "운동" }, external.PushedEvents);
        Assert.Equal(new[] { "운동" }, external.PushedTodos);
    }

    [Fact]
    public async Task TheRoutingMarkerDecidesWhichOutlookFolderIsUsed()
    {
        var (repo, external) = Build();

        await repo.CaptureAsync(Parse("!CD !OL 내일 오후 3시 팀 미팅"), Now);
        await repo.CaptureAsync(Parse("!TD !OL 금요일까지 보고서 제출"), Now);

        Assert.Equal(new[] { "팀 미팅" }, external.PushedEvents);
        Assert.Equal(new[] { "보고서 제출" }, external.PushedTodos);
    }

    [Fact]
    public async Task NothingIsPushedWithoutTheMarker()
    {
        var (repo, external) = Build();

        await repo.CaptureAsync(Parse("내일 오후 9시 운동"), Now);

        Assert.Empty(external.PushedEvents);
        Assert.Empty(external.PushedTodos);
    }

    [Fact]
    public async Task APushedRecordRemembersWhereItWent()
    {
        var (repo, _) = Build();

        var result = await repo.CaptureAsync(Parse("!CD !OL 내일 오후 3시 팀 미팅"), Now);

        // Without this the mirror mode would have nothing to match on and would
        // duplicate the entry on its first read.
        Assert.Equal("fake", result.Event!.ExternalLink!.Provider);
        Assert.Equal("e1", result.Event.ExternalLink.EntryId);
    }

    [Fact]
    public async Task AFailedPushStillKeepsTheEntry()
    {
        var (repo, external) = Build();
        external.ShouldFail = true;

        var result = await repo.CaptureAsync(Parse("!OL 내일 오후 9시 운동"), Now);

        Assert.Single(repo.Todos);
        Assert.Single(repo.Events);
        Assert.Equal("Outlook이 응답하지 않습니다", result.ExternalError);
        Assert.Null(result.Todo!.ExternalLink);
    }

    [Fact]
    public async Task AnAbsentOutlookIsReportedRatherThanThrowing()
    {
        var (repo, external) = Build();
        external.IsAvailable = false;

        var result = await repo.CaptureAsync(Parse("!OL 내일 오후 9시 운동"), Now);

        Assert.Single(repo.Todos);
        Assert.NotNull(result.ExternalError);
        Assert.Null(repo.External);
    }
}

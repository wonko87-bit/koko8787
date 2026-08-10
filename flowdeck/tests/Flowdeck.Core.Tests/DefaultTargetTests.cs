using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Xunit;

namespace Flowdeck.Core.Tests;

public class DefaultTargetTests
{
    // A fixed Thursday, so relative dates in the assertions never drift.
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static ParsedEntry Parse(string input, TargetDefault fallback) =>
        new NaturalLanguageParser(new RoutingRules { DefaultTarget = fallback }).Parse(input, Now);

    [Theory]
    [InlineData(TargetDefault.Todo, EntryTarget.Todo)]
    [InlineData(TargetDefault.Calendar, EntryTarget.Calendar)]
    [InlineData(TargetDefault.Both, EntryTarget.Both)]
    public void AnUnmarkedLineFollowsTheDefault(TargetDefault fallback, EntryTarget expected) =>
        Assert.Equal(expected, Parse("내일 오후 3시 러닝", fallback).Target);

    /// <summary>
    /// The point of naming a default is to stop the guessing. "회의" would otherwise pull
    /// this to the calendar, which is exactly the surprise the setting exists to prevent.
    /// </summary>
    [Fact]
    public void TheDefaultSilencesTheKeywordSweep()
    {
        Assert.Equal(EntryTarget.Todo, Parse("내일 오후 3시 팀 회의", TargetDefault.Todo).Target);
        Assert.Equal(EntryTarget.Calendar, Parse("금요일까지 보고서 제출", TargetDefault.Calendar).Target);
    }

    [Theory]
    [InlineData("!CD 내일 오후 3시 러닝", EntryTarget.Calendar)]
    [InlineData("!BD 내일 오후 3시 러닝", EntryTarget.Both)]
    [InlineData("내일 오후 3시 러닝 #일정", EntryTarget.Calendar)]
    public void AMarkerStillOutranksTheDefault(string input, EntryTarget expected) =>
        Assert.Equal(expected, Parse(input, TargetDefault.Todo).Target);

    [Fact]
    public void AutoIsTheShippedDefaultAndKeepsTheKeywordSweep()
    {
        Assert.Equal(TargetDefault.Auto, new RoutingRules().DefaultTarget);
        Assert.Equal(EntryTarget.Calendar, Parse("내일 오후 3시 팀 회의", TargetDefault.Auto).Target);
        Assert.Equal(EntryTarget.Todo, Parse("금요일까지 보고서 제출", TargetDefault.Auto).Target);
    }

    /// <summary>
    /// A settings file written before this option existed carries only the keyword switch,
    /// so Auto has to keep honouring it rather than quietly turning the sweep back on.
    /// </summary>
    [Fact]
    public void AutoStillHonoursTheOlderKeywordSwitch()
    {
        var rules = new RoutingRules { DefaultTarget = TargetDefault.Auto, UseKeywordHints = false };

        Assert.Equal(EntryTarget.Both, new NaturalLanguageParser(rules).Parse("내일 오후 3시 팀 회의", Now).Target);
    }

    /// <summary>The Outlook flag is on its own axis, so the routing default leaves it alone.</summary>
    [Fact]
    public void TheDefaultTargetDoesNotSendAnythingToOutlook() =>
        Assert.False(Parse("내일 오후 3시 러닝", TargetDefault.Calendar).PushExternal);

    /// <summary>
    /// The pair a heavy Outlook user would set: everything is a todo, everything goes out.
    /// </summary>
    [Fact]
    public void TheTwoDefaultsCombineIntoAStandingMode()
    {
        var parser = new NaturalLanguageParser(new RoutingRules { DefaultTarget = TargetDefault.Todo })
        {
            PushExternalByDefault = true,
        };

        var routine = parser.Parse("금요일까지 보고서 제출", Now);
        Assert.Equal(EntryTarget.Todo, routine.Target);
        Assert.True(routine.PushExternal);

        // ...and each half is still escapable on its own.
        var exception = parser.Parse("!CD !NOL 내일 오후 3시 팀 미팅", Now);
        Assert.Equal(EntryTarget.Calendar, exception.Target);
        Assert.False(exception.PushExternal);
    }
}

public class ExternalSwitchTests
{
    private static readonly DateTime Now = new(2026, 8, 7, 10, 0, 0);

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    [Fact]
    public void TurningTheIntegrationOffLeavesEverythingLocal()
    {
        var external = new FakeExternalStore();
        var repo = new WorkspaceRepository(new InMemoryStore(), external) { ExternalEnabled = false };

        Assert.Null(repo.External);
        Assert.True(repo.ExternalDetected);
    }

    [Fact]
    public async Task AnEntryMarkedForOutlookSaysSoRatherThanFailingSilently()
    {
        var repo = new WorkspaceRepository(new InMemoryStore(), new FakeExternalStore()) { ExternalEnabled = false };

        var result = await repo.CaptureAsync(Parse("!OL 내일 오후 9시 운동"), Now);

        Assert.Single(repo.Todos);
        Assert.Equal("Outlook 연동이 꺼져 있습니다", result.ExternalError);
    }

    /// <summary>A machine with no Outlook is a different message from a switch left off.</summary>
    [Fact]
    public async Task AnAbsentOutlookIsReportedDifferentlyFromAnOffSwitch()
    {
        var external = new FakeExternalStore { IsAvailable = false };
        var repo = new WorkspaceRepository(new InMemoryStore(), external);

        var result = await repo.CaptureAsync(Parse("!OL 내일 오후 9시 운동"), Now);

        Assert.False(repo.ExternalDetected);
        Assert.Equal("Outlook을 찾을 수 없습니다", result.ExternalError);
    }

    [Fact]
    public void WithNoExternalStoreAtAllTheAppIsSimplyLocal()
    {
        var repo = new WorkspaceRepository(new InMemoryStore());

        Assert.False(repo.ExternalDetected);
        Assert.Null(repo.External);
    }

}

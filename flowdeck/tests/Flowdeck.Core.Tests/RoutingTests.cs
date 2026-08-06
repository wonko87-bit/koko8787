using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Xunit;

namespace Flowdeck.Core.Tests;

public class RoutingTests
{
    // A fixed Thursday, so relative dates in the assertions never drift.
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static ParsedEntry Parse(string input, RoutingRules? rules = null) =>
        new NaturalLanguageParser(rules).Parse(input, Now);

    [Fact]
    public void TodoPrefix_RoutesToTodoOnly()
    {
        var entry = Parse("!TD 내일 오후 9시 운동");

        Assert.Equal(EntryTarget.Todo, entry.Target);
        Assert.True(entry.TargetWasExplicit);
        Assert.Equal("운동", entry.Title);
        Assert.Equal(new DateTime(2026, 8, 7, 21, 0, 0), entry.Start);
        Assert.True(entry.HasTime);
    }

    [Fact]
    public void CalendarPrefix_RoutesToCalendarOnly()
    {
        var entry = Parse("!CD 내일 업무보고회의");

        Assert.Equal(EntryTarget.Calendar, entry.Target);
        Assert.True(entry.TargetWasExplicit);
        Assert.Equal("업무보고회의", entry.Title);
        Assert.Equal(new DateTime(2026, 8, 7), entry.Start);
        Assert.True(entry.IsAllDay);
    }

    [Theory]
    [InlineData("!td 내일 운동")]
    [InlineData("/t 내일 운동")]
    [InlineData("!할일 내일 운동")]
    [InlineData("!투두 내일 운동")]
    [InlineData("내일 운동 할일")]
    [InlineData("내일 운동 #할일")]
    public void TodoAliases_AllRouteToTodo(string input) =>
        Assert.Equal(EntryTarget.Todo, Parse(input).Target);

    [Theory]
    [InlineData("!cd 내일 워크샵")]
    [InlineData("/c 내일 워크샵")]
    [InlineData("!일정 내일 워크샵")]
    [InlineData("내일 워크샵 일정")]
    [InlineData("내일 워크샵 #일정")]
    public void CalendarAliases_AllRouteToCalendar(string input) =>
        Assert.Equal(EntryTarget.Calendar, Parse(input).Target);

    [Fact]
    public void NoMarkerAndNoKeyword_GoesToBoth()
    {
        var entry = Parse("내일 오후 9시 운동");

        Assert.Equal(EntryTarget.Both, entry.Target);
        Assert.False(entry.TargetWasExplicit);
        Assert.Equal("운동", entry.Title);
    }

    [Fact]
    public void CalendarKeyword_PullsToCalendar() =>
        Assert.Equal(EntryTarget.Calendar, Parse("내일 오후 3시 팀 회의").Target);

    [Fact]
    public void TodoKeyword_PullsToTodo() =>
        Assert.Equal(EntryTarget.Todo, Parse("금요일까지 보고서 제출").Target);

    [Fact]
    public void KeywordsFromBothLists_FallBackToBoth() =>
        Assert.Equal(EntryTarget.Both, Parse("회의 자료 작성").Target);

    [Fact]
    public void ExplicitMarker_BeatsKeyword() =>
        Assert.Equal(EntryTarget.Todo, Parse("!TD 내일 팀 회의 준비").Target);

    [Fact]
    public void BothMarker_OverridesKeyword() =>
        Assert.Equal(EntryTarget.Both, Parse("!BD 내일 오후 3시 팀 회의").Target);

    [Fact]
    public void KeywordHintsCanBeDisabled()
    {
        var rules = new RoutingRules { UseKeywordHints = false };

        Assert.Equal(EntryTarget.Both, Parse("내일 오후 3시 팀 회의", rules).Target);
        Assert.Equal(EntryTarget.Calendar, Parse("!CD 내일 오후 3시 팀 회의", rules).Target);
    }

    [Fact]
    public void MarkerMustStandAlone_SoOrdinaryWordsSurvive()
    {
        var entry = Parse("일정관리 앱 알아보기");

        Assert.Equal("일정관리 앱 알아보기", entry.Title);
        Assert.False(entry.TargetWasExplicit);
    }

    [Fact]
    public void TrailingMarkerIsStrippedFromTitle()
    {
        var entry = Parse("내일 오후 3시 운동 !TD");

        Assert.Equal(EntryTarget.Todo, entry.Target);
        Assert.Equal("운동", entry.Title);
    }
}

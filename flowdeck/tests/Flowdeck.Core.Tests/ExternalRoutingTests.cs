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

    /// <summary>Entry ids the copy is no longer reachable at — moved or deleted over there.</summary>
    public HashSet<string> Missing { get; } = new();

    public List<string> UpdatedEvents { get; } = new();

    public List<string> UpdatedTodos { get; } = new();

    public List<string> Deleted { get; } = new();

    public Task<bool> UpdateAsync(CalendarEvent calendarEvent, CancellationToken cancellationToken = default)
    {
        if (ShouldFail) throw new InvalidOperationException("Outlook이 응답하지 않습니다");
        if (Missing.Contains(calendarEvent.ExternalLink?.EntryId ?? string.Empty)) return Task.FromResult(false);

        UpdatedEvents.Add(calendarEvent.Title);
        return Task.FromResult(true);
    }

    public Task<bool> UpdateAsync(TodoItem todo, CancellationToken cancellationToken = default)
    {
        if (ShouldFail) throw new InvalidOperationException("Outlook이 응답하지 않습니다");
        if (Missing.Contains(todo.ExternalLink?.EntryId ?? string.Empty)) return Task.FromResult(false);

        UpdatedTodos.Add(todo.Title);
        return Task.FromResult(true);
    }

    public Task<bool> DeleteAsync(ExternalLink link, CancellationToken cancellationToken = default)
    {
        if (ShouldFail) throw new InvalidOperationException("Outlook이 응답하지 않습니다");
        if (Missing.Contains(link.EntryId)) return Task.FromResult(false);

        Deleted.Add(link.EntryId);
        return Task.FromResult(true);
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

/// <summary>
/// Flowdeck is the master and keeps its Outlook copies in step: an entry edited or deleted
/// here is edited or deleted there. Never the other way, and never at the cost of the local
/// change, which is the user's and is already made.
/// </summary>
public class ExternalMirrorTests
{
    private static readonly DateTime Now = new(2026, 8, 7, 10, 0, 0);
    private static readonly DateTime Later = new(2026, 8, 8, 10, 0, 0);

    private static async Task<(WorkspaceRepository Repository, FakeExternalStore Store)> Capture(string line)
    {
        var store = new FakeExternalStore();
        var repository = new WorkspaceRepository(new InMemoryStore(), store);
        await repository.CaptureAsync(new NaturalLanguageParser().Parse(line, Now), Now);
        return (repository, store);
    }

    [Fact]
    public async Task EditingATodoRewritesTheCopy()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");

        var result = await repository.UpdateTodoAsync(
            repository.Todos[0].Id, new EntryEdit { Title = "분기 보고서 제출" }, Later);

        Assert.Equal(ExternalSync.Updated, result.External);
        Assert.Equal(new[] { "분기 보고서 제출" }, store.UpdatedTodos);
    }

    [Fact]
    public async Task EditingAnEventRewritesTheCopy()
    {
        var (repository, store) = await Capture("!OL !CD 내일 오후 3시 기획 회의");

        var result = await repository.UpdateEventAsync(
            repository.Events[0].Id,
            new EntryEdit { Title = "기획 리뷰", When = new DateTime(2026, 8, 12, 15, 0, 0), HasTime = true },
            Later);

        Assert.Equal(ExternalSync.Updated, result.External);
        Assert.Equal(new[] { "기획 리뷰" }, store.UpdatedEvents);
    }

    /// <summary>An entry that was never sent has nothing to keep in step.</summary>
    [Fact]
    public async Task AnEntryWithNoCopyIsLeftAlone()
    {
        var (repository, store) = await Capture("!NOL !TD 내일 오후 3시 보고서 제출");

        var result = await repository.UpdateTodoAsync(
            repository.Todos[0].Id, new EntryEdit { Title = "고침" }, Later);

        Assert.Equal(ExternalSync.None, result.External);
        Assert.Empty(store.UpdatedTodos);
    }

    /// <summary>
    /// Moved between folders in Outlook, or deleted there. Either way the identifier no
    /// longer resolves, and holding the link would make every later edit fail the same way.
    /// </summary>
    [Fact]
    public async Task ACopyThatHasGoneDropsTheLink()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        store.Missing.Add(repository.Todos[0].ExternalLink!.EntryId);

        var result = await repository.UpdateTodoAsync(
            repository.Todos[0].Id, new EntryEdit { Title = "고침" }, Later);

        Assert.Equal(ExternalSync.Lost, result.External);
        Assert.Null(repository.Todos[0].ExternalLink);

        // The local edit still happened. It is the user's, and Outlook does not get a vote.
        Assert.Equal("고침", repository.Todos[0].Title);
    }

    /// <summary>An unreachable mailbox is reported, and changes nothing locally.</summary>
    [Fact]
    public async Task AMailboxThatRefusesDoesNotLoseTheEdit()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        store.ShouldFail = true;

        var result = await repository.UpdateTodoAsync(
            repository.Todos[0].Id, new EntryEdit { Title = "고침" }, Later);

        Assert.Equal(ExternalSync.Failed, result.External);
        Assert.Equal("고침", repository.Todos[0].Title);

        // Still linked: the copy is presumably fine, it just could not be reached.
        Assert.NotNull(repository.Todos[0].ExternalLink);
    }

    /// <summary>The integration being switched off is not the same as an entry having no copy.</summary>
    [Fact]
    public async Task SwitchingTheIntegrationOffStopsTheMirroring()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        repository.ExternalEnabled = false;

        var result = await repository.UpdateTodoAsync(
            repository.Todos[0].Id, new EntryEdit { Title = "고침" }, Later);

        Assert.Equal(ExternalSync.None, result.External);
        Assert.Empty(store.UpdatedTodos);
        Assert.NotNull(repository.Todos[0].ExternalLink);
    }

    [Fact]
    public async Task DeletingTakesTheCopyWithIt()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        var entryId = repository.Todos[0].ExternalLink!.EntryId;

        await repository.DeleteTodoAsync(repository.Todos[0].Id);

        Assert.Equal(new[] { entryId }, store.Deleted);
        Assert.Empty(repository.Todos);
    }

    [Fact]
    public async Task DeletingAnEventTakesTheCopyWithIt()
    {
        var (repository, store) = await Capture("!OL !CD 내일 오후 3시 기획 회의");
        var entryId = repository.Events[0].ExternalLink!.EntryId;

        await repository.DeleteEventAsync(repository.Events[0].Id);

        Assert.Equal(new[] { entryId }, store.Deleted);
        Assert.Empty(repository.Events);
    }

    /// <summary>
    /// The local delete is what the user asked for. A mailbox that will not answer gets a
    /// warning, not a veto — the alternative is an entry that cannot be got rid of.
    /// </summary>
    [Fact]
    public async Task AFailedRemoteDeleteStillDeletesLocallyAndWarns()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        store.ShouldFail = true;

        var warnings = new List<string>();
        repository.ExternalWarning += (_, message) => warnings.Add(message);

        await repository.DeleteTodoAsync(repository.Todos[0].Id);

        Assert.Empty(repository.Todos);
        Assert.Single(warnings);
        Assert.Contains("보고서 제출", warnings[0]);
    }

    /// <summary>Already gone over there is the outcome that was wanted, not a problem.</summary>
    [Fact]
    public async Task DeletingACopyThatIsAlreadyGoneIsQuiet()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        store.Missing.Add(repository.Todos[0].ExternalLink!.EntryId);

        var warnings = new List<string>();
        repository.ExternalWarning += (_, message) => warnings.Add(message);

        await repository.DeleteTodoAsync(repository.Todos[0].Id);

        Assert.Empty(repository.Todos);
        Assert.Empty(warnings);
    }
}

/// <summary>
/// Ticking a todo off here has to tick it off there. The Outlook task folder is what MS To
/// Do and Teams Tasks both show, so one left unfinished shows up unfinished in three places.
/// </summary>
public class ExternalCompletionTests
{
    private static readonly DateTime Now = new(2026, 8, 7, 10, 0, 0);
    private static readonly DateTime Later = new(2026, 8, 8, 10, 0, 0);

    private static async Task<(WorkspaceRepository Repository, FakeExternalStore Store)> Capture(string line)
    {
        var store = new FakeExternalStore();
        var repository = new WorkspaceRepository(new InMemoryStore(), store);
        await repository.CaptureAsync(new NaturalLanguageParser().Parse(line, Now), Now);
        return (repository, store);
    }

    [Fact]
    public async Task TickingOffPushesTheCompletion()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");

        await repository.ToggleTodoAsync(repository.Todos[0].Id, Later);

        Assert.True(repository.Todos[0].IsDone);
        Assert.Equal(new[] { "보고서 제출" }, store.UpdatedTodos);
    }

    /// <summary>Unticking is the same trip, the other way.</summary>
    [Fact]
    public async Task UntickingPushesToo()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        var id = repository.Todos[0].Id;

        await repository.ToggleTodoAsync(id, Later);
        await repository.ToggleTodoAsync(id, Later);

        Assert.False(repository.Todos[0].IsDone);
        Assert.Equal(2, store.UpdatedTodos.Count);
    }

    /// <summary>
    /// A repeating todo is not finished when it is ticked — it moves to its next turn. What
    /// the copy needs is the new date, not a tick, and pushing the whole record gives it
    /// exactly that.
    /// </summary>
    [Fact]
    public async Task ARepeatingTodoPushesItsNextTurnRatherThanACompletion()
    {
        var (repository, store) = await Capture("!OL !TD 매일 오전 7시 스트레칭");
        var before = repository.Todos[0].DueAt;

        await repository.ToggleTodoAsync(repository.Todos[0].Id, Later);

        Assert.False(repository.Todos[0].IsDone);
        Assert.NotEqual(before, repository.Todos[0].DueAt);
        Assert.Single(store.UpdatedTodos);
    }

    [Fact]
    public async Task ATodoWithNoCopyIsTickedOffQuietly()
    {
        var (repository, store) = await Capture("!NOL !TD 내일 오후 3시 보고서 제출");

        await repository.ToggleTodoAsync(repository.Todos[0].Id, Later);

        Assert.True(repository.Todos[0].IsDone);
        Assert.Empty(store.UpdatedTodos);
    }

    /// <summary>The tick is the user's. Outlook not answering does not undo it.</summary>
    [Fact]
    public async Task AFailedPushLeavesTheTickInPlaceAndWarns()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        store.ShouldFail = true;

        var warnings = new List<string>();
        repository.ExternalWarning += (_, message) => warnings.Add(message);

        await repository.ToggleTodoAsync(repository.Todos[0].Id, Later);

        Assert.True(repository.Todos[0].IsDone);
        Assert.NotNull(repository.Todos[0].ExternalLink);
        Assert.Single(warnings);
    }

    [Fact]
    public async Task ACopyThatHasGoneDropsTheLinkAndSaysSo()
    {
        var (repository, store) = await Capture("!OL !TD 내일 오후 3시 보고서 제출");
        store.Missing.Add(repository.Todos[0].ExternalLink!.EntryId);

        var warnings = new List<string>();
        repository.ExternalWarning += (_, message) => warnings.Add(message);

        await repository.ToggleTodoAsync(repository.Todos[0].Id, Later);

        Assert.True(repository.Todos[0].IsDone);
        Assert.Null(repository.Todos[0].ExternalLink);
        Assert.Single(warnings);
    }
}

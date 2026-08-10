using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

public class NoteParsingTests
{
    private static readonly DateTime Now = new(2026, 8, 10, 9, 0, 0);

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    [Fact]
    public void EverythingAfterTheSeparatorIsTheNote()
    {
        var entry = Parse("!TD 금요일까지 보고서 제출 // 3장부터 다시 검토할 것");

        Assert.Equal("보고서 제출", entry.Title);
        Assert.Equal("3장부터 다시 검토할 것", entry.Notes);
    }

    [Fact]
    public void WithoutASeparatorThereIsNoNote() =>
        Assert.Equal(string.Empty, Parse("!TD 금요일까지 보고서 제출").Notes);

    /// <summary>
    /// The note is prose the user wrote for themselves. A date inside it describes something
    /// rather than scheduling it, and must not move the entry.
    /// </summary>
    [Fact]
    public void TheNoteIsNotParsedForDatesOrTags()
    {
        var entry = Parse("!TD 금요일까지 보고서 제출 // 내일 오후 3시에 팀장님이 얘기한 #급함 건");

        Assert.Equal(new DateTime(2026, 8, 14), entry.Start!.Value.Date);
        Assert.Empty(entry.Tags);
        Assert.Equal("내일 오후 3시에 팀장님이 얘기한 #급함 건", entry.Notes);
    }

    [Fact]
    public void BackslashNBreaksTheLine()
    {
        var entry = Parse(@"!TD 장보기 // 우유\n계란\n빵");

        Assert.Equal("우유\n계란\n빵", entry.Notes);
        Assert.Equal(3, entry.Notes.Split('\n').Length);
    }

    /// <summary>
    /// A link is exactly the sort of thing that ends up in a note, and it carries a "//" of
    /// its own. Without the space rule it would cut the line in half.
    /// </summary>
    [Theory]
    [InlineData("!TD 자료 확인 // https://example.com/a//b", "자료 확인", "https://example.com/a//b")]
    [InlineData("!TD https://example.com 열어보기", "https://example.com 열어보기", "")]
    public void ALinkSurvives(string input, string title, string note)
    {
        var entry = Parse(input);

        Assert.Equal(title, entry.Title);
        Assert.Equal(note, entry.Notes);
    }

    [Fact]
    public void OnlyTheFirstSeparatorSplits()
    {
        var entry = Parse("!TD 정리 // 첫째 // 둘째");

        Assert.Equal("정리", entry.Title);
        Assert.Equal("첫째 // 둘째", entry.Notes);
    }

    [Fact]
    public void AnEmptyNoteIsNoNote()
    {
        var entry = Parse("!TD 보고서 제출 //");

        Assert.Equal("보고서 제출", entry.Title);
        Assert.Equal(string.Empty, entry.Notes);
    }

    [Fact]
    public void TheSeparatorCanBeChanged()
    {
        var rules = new RoutingRules { NoteSeparator = ";;" };
        var entry = new NaturalLanguageParser(rules).Parse("!TD 보고서 제출 ;; 메모입니다", Now);

        Assert.Equal("보고서 제출", entry.Title);
        Assert.Equal("메모입니다", entry.Notes);
    }

    /// <summary>Markers are peeled first, so a note does not stop them being found.</summary>
    [Fact]
    public void MarkersStillWorkAlongsideANote()
    {
        var entry = Parse("!CD 내일 오후 3시 팀 회의 #업무 // 안건 정리해서 갈 것");

        Assert.Equal(EntryTarget.Calendar, entry.Target);
        Assert.Equal("팀 회의", entry.Title);
        Assert.Equal(new[] { "업무" }, entry.Tags);
        Assert.Equal("안건 정리해서 갈 것", entry.Notes);
    }
}

public class NoteStorageTests
{
    private static readonly DateTime Now = new(2026, 8, 10, 9, 0, 0);

    private static async Task<WorkspaceRepository> Capture(string line)
    {
        var repository = new WorkspaceRepository(new InMemoryStore());
        await repository.CaptureAsync(new NaturalLanguageParser().Parse(line, Now), Now);
        return repository;
    }

    [Fact]
    public async Task ATodoKeepsItsNote()
    {
        var repository = await Capture("!TD 금요일까지 보고서 제출 // 3장부터");

        Assert.Equal("3장부터", repository.Todos[0].Notes);
    }

    [Fact]
    public async Task AnEventKeepsItsNote()
    {
        var repository = await Capture("!CD 내일 오후 3시 회의 // 안건 정리");

        Assert.Equal("안건 정리", repository.Events[0].Notes);
    }

    /// <summary>One line that becomes two records leaves the same note on both.</summary>
    [Fact]
    public async Task BothHalvesOfABothEntryCarryIt()
    {
        var repository = await Capture("내일 오후 9시 운동 // 하체 위주");

        Assert.Equal("하체 위주", repository.Todos[0].Notes);
        Assert.Equal("하체 위주", repository.Events[0].Notes);
    }
}

public class EntryEditTests
{
    private static readonly DateTime Now = new(2026, 8, 10, 9, 0, 0);
    private static readonly DateTime Later = new(2026, 8, 11, 12, 0, 0);

    private static async Task<WorkspaceRepository> Capture(string line)
    {
        var repository = new WorkspaceRepository(new InMemoryStore());
        await repository.CaptureAsync(new NaturalLanguageParser().Parse(line, Now), Now);
        return repository;
    }

    [Fact]
    public async Task EveryEditableFieldTakes()
    {
        var repository = await Capture("!TD 오늘 오후 3시 보고서 제출");
        var todo = repository.Todos[0];

        var ok = await repository.UpdateTodoAsync(todo.Id, new EntryEdit
        {
            Title = "분기 보고서 제출",
            Notes = "3장부터\n다시",
            When = new DateTime(2026, 8, 14, 17, 0, 0),
            HasTime = true,
            Priority = Priority.Urgent,
            Tags = { "업무", "보고" },
            ReminderMinutesBefore = 60,
        }, Later);

        Assert.True(ok);
        Assert.Equal("분기 보고서 제출", todo.Title);
        Assert.Equal("3장부터\n다시", todo.Notes);
        Assert.Equal(new DateTime(2026, 8, 14, 17, 0, 0), todo.DueAt);
        Assert.Equal(Priority.Urgent, todo.Priority);
        Assert.Equal(new[] { "업무", "보고" }, todo.Tags);
        Assert.Equal(60, todo.ReminderMinutesBefore);
        Assert.Equal(Later, todo.UpdatedAt);
    }

    [Fact]
    public async Task ClearingTheDateTurnsATodoIntoASomedayItem()
    {
        var repository = await Capture("!TD 오늘 오후 3시 보고서 제출 30분 전 알림");

        await repository.UpdateTodoAsync(repository.Todos[0].Id, new EntryEdit
        {
            Title = "보고서 제출",
            When = null,
            ReminderMinutesBefore = 30,
        }, Later);

        Assert.Null(repository.Todos[0].DueAt);

        // A reminder measured from a date that no longer exists cannot fire, so it goes too.
        Assert.Null(repository.Todos[0].ReminderMinutesBefore);
    }

    [Fact]
    public async Task AnEmptyTitleDoesNotLeaveAnUnfindableRow()
    {
        var repository = await Capture("!TD 보고서 제출");

        await repository.UpdateTodoAsync(repository.Todos[0].Id, new EntryEdit { Title = "   " }, Later);

        Assert.Equal(ParsedEntry.UntitledPlaceholder, repository.Todos[0].Title);
    }

    /// <summary>An hour-long meeting dragged to the afternoon is still an hour long.</summary>
    [Fact]
    public async Task MovingAnEventKeepsItsLength()
    {
        var repository = await Capture("!CD 오늘 오후 1시부터 3시까지 워크샵");
        var source = repository.Events[0];

        await repository.UpdateEventAsync(source.Id, new EntryEdit
        {
            Title = "워크샵",
            When = new DateTime(2026, 8, 12, 9, 0, 0),
            HasTime = true,
        }, Later);

        Assert.Equal(new DateTime(2026, 8, 12, 9, 0, 0), source.Start);
        Assert.Equal(new DateTime(2026, 8, 12, 11, 0, 0), source.End);
    }

    [Fact]
    public async Task TurningAnEventAllDayCoversTheWholeDay()
    {
        var repository = await Capture("!CD 오늘 오후 1시 워크샵");
        var source = repository.Events[0];

        await repository.UpdateEventAsync(source.Id, new EntryEdit
        {
            Title = "워크샵",
            When = new DateTime(2026, 8, 12, 1, 0, 0),
            HasTime = false,
        }, Later);

        Assert.True(source.IsAllDay);
        Assert.Equal(new DateTime(2026, 8, 12), source.Start);
        Assert.Equal(new DateTime(2026, 8, 12).AddDays(1).AddSeconds(-1), source.End);
    }

    [Fact]
    public async Task EditingRaisesChangedSoEverythingRedraws()
    {
        var repository = await Capture("!TD 보고서 제출");
        var changed = 0;
        repository.Changed += (_, _) => changed++;

        await repository.UpdateTodoAsync(repository.Todos[0].Id, new EntryEdit { Title = "고침" }, Later);

        Assert.Equal(1, changed);
    }

    [Fact]
    public async Task EditingSomethingAlreadyDeletedIsRefusedRatherThanThrowing()
    {
        var repository = await Capture("!TD 보고서 제출");
        var id = repository.Todos[0].Id;
        await repository.DeleteTodoAsync(id);

        Assert.False(await repository.UpdateTodoAsync(id, new EntryEdit { Title = "고침" }, Later));
        Assert.False(await repository.UpdateEventAsync(id, new EntryEdit { Title = "고침" }, Later));
    }
}

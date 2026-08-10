using Flowdeck.Core.Export;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

public class TransferFileTests
{
    private static readonly DateTimeOffset Exported = new(2026, 8, 10, 21, 0, 0, TimeSpan.FromHours(9));

    private static TodoItem Todo(string title) => new()
    {
        Title = title,
        DueAt = new DateTime(2026, 8, 12, 21, 0, 0),
        HasTime = true,
        Priority = Priority.High,
        Tags = { "아이디어", "업무" },
        Recurrence = new Recurrence { Kind = RecurrenceKind.Weekly, Interval = 1, DaysOfWeek = { DayOfWeek.Monday } },
        ReminderMinutesBefore = 30,
        SourceText = "!TD 매주 월요일 오후 9시 회고 #아이디어 #업무",
        Notes = "줄바꿈\n포함",
    };

    [Fact]
    public void EverythingThatMattersSurvivesTheRoundTrip()
    {
        var original = Todo("주간 회고");

        var archive = TransferFile.Read(TransferFile.Write(new[] { original }, Array.Empty<CalendarEvent>(), Exported));
        var copy = Assert.Single(archive.Todos);

        Assert.Equal(original.Id, copy.Id);
        Assert.Equal("주간 회고", copy.Title);
        Assert.Equal(original.DueAt, copy.DueAt);
        Assert.True(copy.HasTime);
        Assert.Equal(Priority.High, copy.Priority);
        Assert.Equal(new[] { "아이디어", "업무" }, copy.Tags);
        Assert.Equal(RecurrenceKind.Weekly, copy.Recurrence.Kind);
        Assert.Equal(new[] { DayOfWeek.Monday }, copy.Recurrence.DaysOfWeek);
        Assert.Equal(30, copy.ReminderMinutesBefore);
        Assert.Equal(original.SourceText, copy.SourceText);
        Assert.Equal("줄바꿈\n포함", copy.Notes);
    }

    [Fact]
    public void EventsRoundTripToo()
    {
        var source = new CalendarEvent
        {
            Title = "워크샵",
            Start = new DateTime(2026, 8, 12),
            End = new DateTime(2026, 8, 14),
            IsAllDay = true,
            Tags = { "행사" },
        };

        var archive = TransferFile.Read(TransferFile.Write(Array.Empty<TodoItem>(), new[] { source }, Exported));
        var copy = Assert.Single(archive.Events);

        Assert.Equal(source.Id, copy.Id);
        Assert.Equal("워크샵", copy.Title);
        Assert.True(copy.IsAllDay);
        Assert.Equal(source.Start, copy.Start);
        Assert.Equal(new[] { "행사" }, copy.Tags);
    }

    /// <summary>
    /// A link names an entry inside one machine's Outlook. Carried across, the receiving
    /// copy would look like it had already been sent — and the push guard would make sure
    /// it never was.
    /// </summary>
    [Fact]
    public void TheExternalLinkDoesNotTravel()
    {
        var todo = Todo("주간 회고");
        todo.ExternalLink = new ExternalLink { Provider = "outlook", EntryId = "0000A" };

        var archive = TransferFile.Read(TransferFile.Write(new[] { todo }, Array.Empty<CalendarEvent>(), Exported));

        Assert.Null(Assert.Single(archive.Todos).ExternalLink);

        // ...and the record still being used by the exporting app keeps its own link.
        Assert.NotNull(todo.ExternalLink);
    }

    [Fact]
    public void TheFileNamesItsOwnShape()
    {
        var json = TransferFile.Write(new[] { Todo("x") }, Array.Empty<CalendarEvent>(), Exported);

        Assert.Contains(TransferArchive.FormatName, json);
        Assert.Contains("\"Version\": 1", json);
    }

    [Theory]
    [InlineData("not json at all")]
    [InlineData("{ \"Format\": \"something.else\", \"Version\": 1 }")]
    [InlineData("{ \"Format\": \"flowdeck.transfer\", \"Version\": 99 }")]
    public void AFileThatIsNotOneIsRefusedWithAReadableReason(string json)
    {
        var error = Assert.Throws<FormatException>(() => TransferFile.Read(json));

        Assert.False(string.IsNullOrWhiteSpace(error.Message));
        Assert.DoesNotContain("Exception", error.Message);
    }

    [Fact]
    public void AFileWithNullListsReadsAsEmptyRatherThanExploding()
    {
        var archive = TransferFile.Read(
            "{ \"Format\": \"flowdeck.transfer\", \"Version\": 1, \"Todos\": null, \"Events\": null }");

        Assert.Empty(archive.Todos);
        Assert.Empty(archive.Events);
    }

    [Fact]
    public void AnEmptyExportIsStillAValidFile()
    {
        var archive = TransferFile.Read(
            TransferFile.Write(Array.Empty<TodoItem>(), Array.Empty<CalendarEvent>(), Exported));

        Assert.Equal(0, archive.Count);
    }
}

public class ImportTests
{
    private static readonly DateTime Now = new(2026, 8, 10, 10, 0, 0);
    private static readonly DateTimeOffset Exported = new(2026, 8, 10, 21, 0, 0, TimeSpan.FromHours(9));

    private static async Task<WorkspaceRepository> Populated(params string[] lines)
    {
        var repo = new WorkspaceRepository(new InMemoryStore());
        var parser = new NaturalLanguageParser();

        foreach (var line in lines) await repo.CaptureAsync(parser.Parse(line, Now), Now);

        return repo;
    }

    private static TransferArchive Carry(WorkspaceRepository from) =>
        TransferFile.Read(TransferFile.Write(from.Todos, from.Events, Exported));

    [Fact]
    public async Task ImportingAddsTheEntries()
    {
        var source = await Populated("!TD 내일 아이디어 정리", "!CD 8월 20일 오후 3시 검토 회의");
        var target = new WorkspaceRepository(new InMemoryStore());

        var result = await target.ImportAsync(Carry(source));

        Assert.Equal(1, result.TodosAdded);
        Assert.Equal(1, result.EventsAdded);
        Assert.Equal(0, result.Skipped);
        Assert.Equal("아이디어 정리", target.Todos[0].Title);
    }

    /// <summary>Carrying the same file in twice is something people do. It must be harmless.</summary>
    [Fact]
    public async Task ImportingTheSameFileTwiceChangesNothingTheSecondTime()
    {
        var source = await Populated("!TD 내일 아이디어 정리");
        var target = new WorkspaceRepository(new InMemoryStore());
        var archive = Carry(source);

        await target.ImportAsync(archive);
        var second = await target.ImportAsync(Carry(source));

        Assert.Single(target.Todos);
        Assert.Equal(0, second.Added);
        Assert.Equal(1, second.Skipped);
    }

    /// <summary>
    /// The file travels one way and an entry already here has probably been worked on since
    /// it arrived. Re-importing must not roll that back.
    /// </summary>
    [Fact]
    public async Task AnEntryEditedOnThisMachineIsNotOverwritten()
    {
        var source = await Populated("!TD 내일 아이디어 정리");
        var target = new WorkspaceRepository(new InMemoryStore());
        await target.ImportAsync(Carry(source));

        target.Todos[0].Title = "여기서 다듬은 제목";
        await target.ImportAsync(Carry(source));

        Assert.Equal("여기서 다듬은 제목", target.Todos[0].Title);
    }

    [Fact]
    public async Task ImportingNeverRemovesWhatIsAlreadyHere()
    {
        var source = await Populated("!TD 내일 아이디어 정리");
        var target = await Populated("!TD 여기서 만든 할일");

        await target.ImportAsync(Carry(source));

        Assert.Equal(2, target.Todos.Count);
        Assert.Contains(target.Todos, t => t.Title == "여기서 만든 할일");
    }

    [Fact]
    public async Task AnImportedRecordArrivesUnsent()
    {
        var source = await Populated("!TD 내일 아이디어 정리");
        source.Todos[0].ExternalLink = new ExternalLink { Provider = "outlook", EntryId = "0000A" };

        var target = new WorkspaceRepository(new InMemoryStore());
        await target.ImportAsync(Carry(source));

        Assert.Null(target.Todos[0].ExternalLink);
    }

    /// <summary>
    /// The list windows show one kind at a time, so exporting half of a both-targets pair is
    /// ordinary. The half that arrives must not hold the id of something that never came.
    /// </summary>
    [Fact]
    public async Task AHalfExportedPairArrivesWithoutADanglingLink()
    {
        var source = await Populated("내일 오후 9시 운동");
        Assert.NotNull(source.Todos[0].LinkedEventId);

        var todosOnly = TransferFile.Read(
            TransferFile.Write(source.Todos, Array.Empty<CalendarEvent>(), Exported));

        var target = new WorkspaceRepository(new InMemoryStore());
        await target.ImportAsync(todosOnly);

        Assert.Single(target.Todos);
        Assert.Null(target.Todos[0].LinkedEventId);
    }

    [Fact]
    public async Task AWholePairKeepsItsLink()
    {
        var source = await Populated("내일 오후 9시 운동");
        var target = new WorkspaceRepository(new InMemoryStore());

        await target.ImportAsync(Carry(source));

        Assert.Equal(target.Events[0].Id, target.Todos[0].LinkedEventId);
        Assert.Equal(target.Todos[0].Id, target.Events[0].LinkedTodoId);
    }

    [Fact]
    public async Task AnEmptyImportRaisesNothing()
    {
        var target = new WorkspaceRepository(new InMemoryStore());
        var changed = 0;
        target.Changed += (_, _) => changed++;

        var result = await target.ImportAsync(new TransferArchive());

        Assert.Equal(0, result.Added);
        Assert.Equal(0, changed);
        Assert.Contains("없습니다", result.Describe());
    }

    [Fact]
    public async Task ImportedEntriesShowUpInTheOrdinaryViews()
    {
        var source = await Populated("!CD 8월 20일 오후 3시 검토 회의");
        var target = new WorkspaceRepository(new InMemoryStore());

        await target.ImportAsync(Carry(source));

        Assert.Single(target.OccurrencesOn(new DateTime(2026, 8, 20)));
    }
}

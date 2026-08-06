using Flowdeck.Core.Export;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>Keeps the workspace in memory so repository behaviour can be tested without disk.</summary>
internal sealed class InMemoryStore : IWorkspaceStore
{
    private Workspace _workspace = new();

    public int SaveCount { get; private set; }

    public Task<Workspace> LoadAsync(CancellationToken cancellationToken = default) =>
        Task.FromResult(_workspace);

    public Task SaveAsync(Workspace workspace, CancellationToken cancellationToken = default)
    {
        _workspace = workspace;
        SaveCount++;
        return Task.CompletedTask;
    }
}

public class EntryComposerTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    [Fact]
    public void BothTarget_CreatesLinkedRecords()
    {
        var result = EntryComposer.Compose(Parse("내일 오후 9시 운동"), Now);

        Assert.NotNull(result.Todo);
        Assert.NotNull(result.Event);
        Assert.Equal(result.Event!.Id, result.Todo!.LinkedEventId);
        Assert.Equal(result.Todo.Id, result.Event.LinkedTodoId);
        Assert.Equal("운동", result.Todo.Title);
    }

    [Fact]
    public void BothTargetWithoutADate_StaysATodo()
    {
        var result = EntryComposer.Compose(Parse("영수증 정리"), Now);

        Assert.NotNull(result.Todo);
        Assert.Null(result.Event);
        Assert.Null(result.Todo!.DueAt);
    }

    [Fact]
    public void TodoMarker_CreatesNoEvent()
    {
        var result = EntryComposer.Compose(Parse("!TD 내일 오후 9시 운동"), Now);

        Assert.NotNull(result.Todo);
        Assert.Null(result.Event);
    }

    [Fact]
    public void CalendarMarker_CreatesNoTodo()
    {
        var result = EntryComposer.Compose(Parse("!CD 내일 업무보고회의"), Now);

        Assert.Null(result.Todo);
        Assert.NotNull(result.Event);
        Assert.True(result.Event!.IsAllDay);
        Assert.Equal(new DateTime(2026, 8, 7), result.Event.Start);
        Assert.Equal(new DateTime(2026, 8, 7, 23, 59, 59), result.Event.End);
    }
}

public class WorkspaceRepositoryTests
{
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static (WorkspaceRepository Repo, InMemoryStore Store) Build()
    {
        var store = new InMemoryStore();
        return (new WorkspaceRepository(store), store);
    }

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    [Fact]
    public async Task CapturePersistsBothRecords()
    {
        var (repo, store) = Build();

        await repo.CaptureAsync(Parse("내일 오후 9시 운동"), Now);

        Assert.Single(repo.Todos);
        Assert.Single(repo.Events);
        Assert.Equal(1, store.SaveCount);
    }

    [Fact]
    public async Task TogglingAPlainTodoCompletesIt()
    {
        var (repo, _) = Build();
        await repo.CaptureAsync(Parse("!TD 영수증 정리"), Now);
        var id = repo.Todos[0].Id;

        await repo.ToggleTodoAsync(id, Now);

        Assert.True(repo.Todos[0].IsDone);
        Assert.Equal(Now, repo.Todos[0].CompletedAt);
    }

    [Fact]
    public async Task TogglingARepeatingTodoAdvancesItInsteadOfClosingIt()
    {
        var (repo, _) = Build();
        await repo.CaptureAsync(Parse("!TD 매일 오전 7시 스트레칭"), Now);
        var todo = repo.Todos[0];
        var firstDue = todo.DueAt;

        await repo.ToggleTodoAsync(todo.Id, Now);

        Assert.False(repo.Todos[0].IsDone);
        Assert.Equal(firstDue!.Value.AddDays(1), repo.Todos[0].DueAt);
    }

    [Fact]
    public async Task DeletingATodoClearsTheLinkOnItsEvent()
    {
        var (repo, _) = Build();
        await repo.CaptureAsync(Parse("내일 오후 9시 운동"), Now);
        var todoId = repo.Todos[0].Id;

        await repo.DeleteTodoAsync(todoId);

        Assert.Empty(repo.Todos);
        Assert.Single(repo.Events);
        Assert.Null(repo.Events[0].LinkedTodoId);
    }

    [Fact]
    public async Task OccurrencesOnExpandsRepeatingEvents()
    {
        var (repo, _) = Build();
        await repo.CaptureAsync(Parse("!CD 매주 화요일 오전 10시 주간회의"), Now);

        Assert.Single(repo.OccurrencesOn(new DateTime(2026, 8, 11)));
        Assert.Single(repo.OccurrencesOn(new DateTime(2026, 8, 18)));
        Assert.Empty(repo.OccurrencesOn(new DateTime(2026, 8, 12)));
    }

    [Fact]
    public async Task TodosForTodayPutsOverdueItemsFirst()
    {
        var (repo, _) = Build();
        await repo.CaptureAsync(Parse("!TD 어제 오전 9시 밀린 보고서"), Now);
        await repo.CaptureAsync(Parse("!TD 오늘 오후 3시 팀 공유"), Now);
        await repo.CaptureAsync(Parse("!TD 언젠가 책 읽기"), Now);

        var list = repo.TodosForToday(Now);

        Assert.Equal(3, list.Count);
        Assert.Equal("밀린 보고서", list[0].Title);
        Assert.Equal("팀 공유", list[1].Title);
        Assert.Null(list[2].DueAt);
    }

    [Fact]
    public async Task PurgeRemovesOldCompletedItems()
    {
        var (repo, _) = Build();
        await repo.CaptureAsync(Parse("!TD 영수증 정리"), Now);
        await repo.ToggleTodoAsync(repo.Todos[0].Id, Now.AddDays(-30));

        var removed = await repo.PurgeCompletedAsync(Now.AddDays(-7));

        Assert.Equal(1, removed);
        Assert.Empty(repo.Todos);
    }
}

public class JsonWorkspaceStoreTests
{
    [Fact]
    public async Task RoundTripsThroughDisk()
    {
        var directory = Path.Combine(Path.GetTempPath(), "flowdeck-tests-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "workspace.json");

        try
        {
            var store = new JsonWorkspaceStore(path);
            var workspace = new Workspace();
            workspace.Todos.Add(new TodoItem
            {
                Title = "운동",
                DueAt = new DateTime(2026, 8, 7, 21, 0, 0),
                HasTime = true,
                Priority = Priority.High,
                Tags = { "건강" },
                Recurrence = new Recurrence { Kind = RecurrenceKind.Weekly, DaysOfWeek = { DayOfWeek.Friday } },
            });
            workspace.Events.Add(new CalendarEvent
            {
                Title = "업무보고회의",
                Start = new DateTime(2026, 8, 7),
                End = new DateTime(2026, 8, 7, 23, 59, 59),
                IsAllDay = true,
            });

            await store.SaveAsync(workspace);
            var reloaded = await store.LoadAsync();

            Assert.Equal("운동", reloaded.Todos[0].Title);
            Assert.Equal(Priority.High, reloaded.Todos[0].Priority);
            Assert.Equal(new[] { DayOfWeek.Friday }, reloaded.Todos[0].Recurrence.DaysOfWeek);
            Assert.Equal("업무보고회의", reloaded.Events[0].Title);
            Assert.True(reloaded.Events[0].IsAllDay);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task MissingFileLoadsAnEmptyWorkspace()
    {
        var path = Path.Combine(Path.GetTempPath(), "flowdeck-missing-" + Guid.NewGuid().ToString("N"), "w.json");

        var workspace = await new JsonWorkspaceStore(path).LoadAsync();

        Assert.Empty(workspace.Todos);
        Assert.Empty(workspace.Events);
    }

    [Fact]
    public async Task CorruptFileIsSetAsideRatherThanThrowing()
    {
        var directory = Path.Combine(Path.GetTempPath(), "flowdeck-corrupt-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "workspace.json");

        try
        {
            await File.WriteAllTextAsync(path, "{ this is not json");

            var workspace = await new JsonWorkspaceStore(path).LoadAsync();

            Assert.Empty(workspace.Todos);
            Assert.True(File.Exists(path + ".corrupt"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}

public class IcsExporterTests
{
    private static readonly DateTime Stamp = new(2026, 8, 6, 10, 0, 0, DateTimeKind.Utc);

    [Fact]
    public void TimedEventCarriesStartAndEnd()
    {
        var ics = IcsExporter.Export(
            new[]
            {
                new CalendarEvent
                {
                    Title = "팀 회의",
                    Start = new DateTime(2026, 8, 7, 15, 0, 0),
                    End = new DateTime(2026, 8, 7, 16, 0, 0),
                },
            },
            Stamp);

        Assert.Contains("BEGIN:VEVENT", ics);
        Assert.Contains("DTSTART:20260807T150000", ics);
        Assert.Contains("DTEND:20260807T160000", ics);
        Assert.Contains("SUMMARY:팀 회의", ics);
    }

    [Fact]
    public void AllDayEventEndsOnTheFollowingDate()
    {
        var ics = IcsExporter.Export(
            new[]
            {
                new CalendarEvent
                {
                    Title = "워크샵",
                    Start = new DateTime(2026, 8, 10),
                    End = new DateTime(2026, 8, 12, 23, 59, 59),
                    IsAllDay = true,
                },
            },
            Stamp);

        Assert.Contains("DTSTART;VALUE=DATE:20260810", ics);
        Assert.Contains("DTEND;VALUE=DATE:20260813", ics);
    }

    [Fact]
    public void WeeklyRuleBecomesAnRrule()
    {
        var ics = IcsExporter.Export(
            new[]
            {
                new CalendarEvent
                {
                    Title = "주간회의",
                    Start = new DateTime(2026, 8, 11, 10, 0, 0),
                    End = new DateTime(2026, 8, 11, 11, 0, 0),
                    Recurrence = new Recurrence
                    {
                        Kind = RecurrenceKind.Weekly,
                        Interval = 2,
                        DaysOfWeek = { DayOfWeek.Tuesday },
                    },
                },
            },
            Stamp);

        Assert.Contains("RRULE:FREQ=WEEKLY;BYDAY=TU;INTERVAL=2", ics);
    }

    [Fact]
    public void ReminderBecomesAValarm()
    {
        var ics = IcsExporter.Export(
            new[]
            {
                new CalendarEvent
                {
                    Title = "발표",
                    Start = new DateTime(2026, 8, 11, 10, 0, 0),
                    End = new DateTime(2026, 8, 11, 11, 0, 0),
                    ReminderMinutesBefore = 30,
                },
            },
            Stamp);

        Assert.Contains("TRIGGER:-PT30M", ics);
    }

    [Fact]
    public void StructuralCharactersAreEscaped()
    {
        var ics = IcsExporter.Export(
            new[]
            {
                new CalendarEvent
                {
                    Title = "예산, 인력 검토; 1차",
                    Start = new DateTime(2026, 8, 11, 10, 0, 0),
                    End = new DateTime(2026, 8, 11, 11, 0, 0),
                },
            },
            Stamp);

        Assert.Contains(@"SUMMARY:예산\, 인력 검토\; 1차", ics);
    }
}

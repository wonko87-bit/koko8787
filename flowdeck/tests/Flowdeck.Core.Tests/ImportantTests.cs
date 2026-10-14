using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// The few things that matter most, kept in sight whichever day is being looked at: todos
/// and events at High or above, and where an event gets its priority from.
/// </summary>
public class ImportantTests
{
    private static readonly DateTime Now = new(2026, 10, 14, 9, 0, 0);

    private static async Task<WorkspaceRepository> Repository(Workspace workspace)
    {
        var repository = new WorkspaceRepository(new InMemoryStore());
        await repository.ReplaceAsync(workspace);
        return repository;
    }

    private static TodoItem Todo(string title, Priority priority, DateTime? due = null, bool done = false) =>
        new() { Title = title, Priority = priority, DueAt = due, HasTime = due.HasValue, IsDone = done };

    private static CalendarEvent Event(string title, Priority priority, DateTime start, Recurrence? recurrence = null) =>
        new() { Title = title, Priority = priority, Start = start, End = start.AddHours(1), Recurrence = recurrence ?? Recurrence.None };

    // ---- what counts --------------------------------------------------------

    [Theory]
    [InlineData(Priority.Urgent, true)]
    [InlineData(Priority.High, true)]
    [InlineData(Priority.Normal, false)]
    [InlineData(Priority.Low, false)]
    [InlineData(Priority.None, false)]
    public void HighAndAboveIsImportant(Priority priority, bool important) =>
        Assert.Equal(important, priority.IsImportant());

    // ---- todos ----------------------------------------------------------------

    [Fact]
    public async Task ImportantTodosAreTheOpenOnesAtHighOrAboveSoonestFirstUndatedLast()
    {
        var workspace = new Workspace();
        workspace.Todos.Add(Todo("언젠가 급함", Priority.Urgent));
        workspace.Todos.Add(Todo("보통", Priority.Normal, Now.AddDays(1)));
        workspace.Todos.Add(Todo("다음주 중요", Priority.High, Now.AddDays(7)));
        workspace.Todos.Add(Todo("내일 급함", Priority.Urgent, Now.AddDays(1)));
        workspace.Todos.Add(Todo("끝난 급함", Priority.Urgent, Now.AddDays(2), done: true));
        var repository = await Repository(workspace);

        Assert.Equal(new[] { "내일 급함", "다음주 중요", "언젠가 급함" }, repository.ImportantTodos().Select(t => t.Title));
    }

    [Fact]
    public async Task AnOverdueImportantTodoStaysInTheList()
    {
        // Overdue is exactly when it matters most.
        var workspace = new Workspace();
        workspace.Todos.Add(Todo("지난 급함", Priority.Urgent, Now.AddDays(-3)));
        var repository = await Repository(workspace);

        Assert.Single(repository.ImportantTodos());
    }

    // ---- events ---------------------------------------------------------------

    [Fact]
    public async Task ImportantEventsAreTheOnesAheadAtHighOrAboveSoonestFirst()
    {
        var workspace = new Workspace();
        workspace.Events.Add(Event("지난 중요", Priority.High, Now.AddDays(-2)));
        workspace.Events.Add(Event("보통 회의", Priority.Normal, Now.AddDays(1)));
        workspace.Events.Add(Event("다음주 중요", Priority.High, Now.AddDays(7)));
        workspace.Events.Add(Event("내일 급함", Priority.Urgent, Now.AddDays(1)));
        var repository = await Repository(workspace);

        Assert.Equal(new[] { "내일 급함", "다음주 중요" }, repository.ImportantOccurrences(Now).Select(o => o.Title));
    }

    [Fact]
    public async Task AnImportantMeetingUnderWayRightNowStillCounts()
    {
        var workspace = new Workspace();
        workspace.Events.Add(Event("진행 중", Priority.High, Now.AddMinutes(-20)));
        var repository = await Repository(workspace);

        Assert.Single(repository.ImportantOccurrences(Now));
    }

    [Fact]
    public async Task ARepeatingImportantEventContributesItsNextOccurrenceOnly()
    {
        var weekly = new Recurrence { Kind = RecurrenceKind.Weekly, Interval = 1, DaysOfWeek = { DayOfWeek.Monday } };
        var workspace = new Workspace();
        workspace.Events.Add(Event("주간 리뷰", Priority.High, new DateTime(2026, 9, 7, 10, 0, 0), weekly));
        var repository = await Repository(workspace);

        var occurrence = Assert.Single(repository.ImportantOccurrences(Now));
        Assert.Equal(new DateTime(2026, 10, 19, 10, 0, 0), occurrence.Start);
    }

    // ---- where an event's priority comes from ------------------------------------

    [Fact]
    public void TheWordInTheLineSetsTheEventsPriority()
    {
        var entry = new NaturalLanguageParser().Parse("!CD 중요 설계 리뷰 내일 2시", Now);

        var result = EntryComposer.Compose(entry, Now);

        Assert.Equal(Priority.High, result.Event!.Priority);
    }

    [Fact]
    public void ANumberMarkerSetsItToo()
    {
        var entry = new NaturalLanguageParser().Parse("!CD !1 임원 보고 내일 2시", Now);

        Assert.Equal(Priority.Urgent, EntryComposer.Compose(entry, Now).Event!.Priority);
    }

    [Fact]
    public async Task EditingAnEventKeepsAndChangesItsPriority()
    {
        var workspace = new Workspace();
        var meeting = Event("리뷰", Priority.None, Now.AddDays(1));
        workspace.Events.Add(meeting);
        var repository = await Repository(workspace);

        await repository.UpdateEventAsync(meeting.Id, new EntryEdit { Title = "리뷰", When = meeting.Start, HasTime = true, Priority = Priority.Urgent }, Now);

        Assert.Equal(Priority.Urgent, meeting.Priority);
        Assert.Single(repository.ImportantOccurrences(Now));
    }

    [Fact]
    public void AnEventsPriorityTravelsInATransferFile()
    {
        var meeting = Event("중요 회의", Priority.High, Now.AddDays(1));

        var text = Flowdeck.Core.Export.TransferFile.Write(Array.Empty<TodoItem>(), new[] { meeting }, DateTimeOffset.Now);

        Assert.Equal(Priority.High, Assert.Single(Flowdeck.Core.Export.TransferFile.Read(text).Events).Priority);
    }
}

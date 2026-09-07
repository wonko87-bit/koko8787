using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// Todos that belong to an event: made from the event, due when it is, moving with it, and
/// outliving it.
/// </summary>
public class AttachedTodoTests
{
    private static readonly DateTime Now = new(2026, 8, 17, 9, 0, 0);
    private static readonly DateTime Thursday = new(2026, 8, 20, 14, 0, 0);

    private static async Task<(WorkspaceRepository Repository, CalendarEvent Meeting)> Setup(bool allDay = false)
    {
        var meeting = new CalendarEvent
        {
            Title = "설계 리뷰",
            Start = allDay ? Thursday.Date : Thursday,
            End = allDay ? Thursday.Date.AddDays(1).AddSeconds(-1) : Thursday.AddHours(1),
            IsAllDay = allDay,
            Tags = { "모터A" },
        };

        var workspace = new Workspace();
        workspace.Events.Add(meeting);

        var repository = new WorkspaceRepository(new InMemoryStore());
        await repository.ReplaceAsync(workspace);
        return (repository, meeting);
    }

    [Fact]
    public async Task ATodoMadeFromAnEventIsDueWhenTheEventIsAndCarriesItsTags()
    {
        var (repository, meeting) = await Setup();

        var todo = await repository.AttachTodoAsync(meeting.Id, "발표 자료 작성", Now);

        Assert.Equal("발표 자료 작성", todo.Title);
        Assert.Equal(Thursday, todo.DueAt);
        Assert.True(todo.HasTime);
        Assert.Equal(new[] { "모터A" }, todo.Tags);
        Assert.Equal(meeting.Id, todo.LinkedEventId);
        Assert.False(todo.IsDone);
        Assert.Contains(todo, repository.Todos);
    }

    [Fact]
    public async Task AnAllDayEventGivesADateWithoutAClock()
    {
        var (repository, meeting) = await Setup(allDay: true);

        var todo = await repository.AttachTodoAsync(meeting.Id, "준비", Now);

        Assert.Equal(Thursday.Date, todo.DueAt);
        Assert.False(todo.HasTime);
    }

    [Fact]
    public async Task AnEventThatIsGoneCannotHaveThingsAttached()
    {
        var (repository, _) = await Setup();

        await Assert.ThrowsAsync<InvalidOperationException>(() => repository.AttachTodoAsync("없는-id", "x", Now));
    }

    [Fact]
    public async Task WhatIsAttachedComesBackOpenFirstThenByDate()
    {
        var (repository, meeting) = await Setup();
        var later = await repository.AttachTodoAsync(meeting.Id, "나중", Now);
        var done = await repository.AttachTodoAsync(meeting.Id, "끝난 것", Now);
        var sooner = await repository.AttachTodoAsync(meeting.Id, "먼저", Now);

        await repository.UpdateTodoAsync(sooner.Id, new EntryEdit { Title = "먼저", When = Thursday.AddDays(-1), HasTime = true }, Now);
        await repository.ToggleTodoAsync(done.Id, Now);

        Assert.Equal(new[] { "먼저", "나중", "끝난 것" }, repository.TodosAttachedTo(meeting.Id).Select(t => t.Title));
    }

    [Fact]
    public async Task AnUnrelatedTodoIsNotAttached()
    {
        var (repository, meeting) = await Setup();
        await repository.AttachTodoAsync(meeting.Id, "붙은 것", Now);
        await repository.CaptureAsync(new NaturalLanguageParser().Parse("!TD 상관없는 것", Now), Now);

        Assert.Equal("붙은 것", Assert.Single(repository.TodosAttachedTo(meeting.Id)).Title);
    }

    [Fact]
    public async Task DetachingKeepsTheTodoAndCutsTheTieBothWays()
    {
        var (repository, meeting) = await Setup();
        var todo = await repository.AttachTodoAsync(meeting.Id, "자료", Now);
        meeting.LinkedTodoId = todo.Id;

        await repository.DetachTodoAsync(todo.Id);

        Assert.Null(todo.LinkedEventId);
        Assert.Null(meeting.LinkedTodoId);
        Assert.Contains(todo, repository.Todos);
        Assert.Empty(repository.TodosAttachedTo(meeting.Id));
    }

    [Fact]
    public async Task DeletingTheEventLeavesItsTodosBehindUntied()
    {
        var (repository, meeting) = await Setup();
        var todo = await repository.AttachTodoAsync(meeting.Id, "자료", Now);

        await repository.DeleteEventAsync(meeting.Id);

        Assert.Contains(todo, repository.Todos);
        Assert.Null(todo.LinkedEventId);
    }

    [Fact]
    public async Task MovingTheEventMovesWhatIsAttachedAndOpen()
    {
        var (repository, meeting) = await Setup();
        var slides = await repository.AttachTodoAsync(meeting.Id, "자료", Now);
        var room = await repository.AttachTodoAsync(meeting.Id, "회의실", Now);
        await repository.ToggleTodoAsync(room.Id, Now);

        await repository.UpdateEventAsync(meeting.Id, new EntryEdit { Title = meeting.Title, When = Thursday.AddDays(5), HasTime = true }, Now);

        Assert.Equal(Thursday.AddDays(5), slides.DueAt);
        Assert.Equal(Thursday, room.DueAt);
    }

    [Fact]
    public async Task ATwinMadeByBothIsAttachedToItsEvent()
    {
        // "둘 다" makes an event and a todo that point at each other; the todo shows up
        // among the event's attachments like any other, because it is one.
        var repository = new WorkspaceRepository(new InMemoryStore());
        await repository.LoadAsync();

        var result = await repository.CaptureAsync(new NaturalLanguageParser().Parse("!BD 내일 3시 리뷰", Now), Now);

        Assert.Equal(result.Todo!.Id, Assert.Single(repository.TodosAttachedTo(result.Event!.Id)).Id);
    }
}

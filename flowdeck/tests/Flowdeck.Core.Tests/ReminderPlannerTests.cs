using Flowdeck.Core.Parsing;
using Flowdeck.Core.Services;
using Flowdeck.Core.Storage;
using Xunit;

namespace Flowdeck.Core.Tests;

public class ReminderPlannerTests
{
    // A fixed Monday morning.
    private static readonly DateTime Now = new(2026, 8, 10, 9, 0, 0);

    private static async Task<WorkspaceRepository> Repo(params string[] lines)
    {
        var repository = new WorkspaceRepository(new InMemoryStore());
        var parser = new NaturalLanguageParser();

        foreach (var line in lines) await repository.CaptureAsync(parser.Parse(line, Now), Now);

        return repository;
    }

    private static IReadOnlyList<DueReminder> Upcoming(WorkspaceRepository repository, TimeSpan? horizon = null) =>
        ReminderPlanner.Upcoming(repository, Now, horizon ?? TimeSpan.FromDays(7));

    [Fact]
    public async Task ATodoReminderLandsBeforeItsDueTime()
    {
        var repository = await Repo("!TD 오늘 오후 3시 보고서 제출 30분 전 알림");

        var reminder = Assert.Single(Upcoming(repository));

        Assert.Equal(new DateTime(2026, 8, 10, 14, 30, 0), reminder.At);
        Assert.Equal(new DateTime(2026, 8, 10, 15, 0, 0), reminder.Subject);
        Assert.True(reminder.IsTodo);
        Assert.Equal("보고서 제출", reminder.Title);
    }

    [Fact]
    public async Task AnEventReminderLandsBeforeItsStart()
    {
        var repository = await Repo("!CD 오늘 오후 3시 팀 회의 1시간 전 알림");

        var reminder = Assert.Single(Upcoming(repository));

        Assert.Equal(new DateTime(2026, 8, 10, 14, 0, 0), reminder.At);
        Assert.False(reminder.IsTodo);
    }

    [Fact]
    public async Task AnEntryWithoutAReminderIsSilent() =>
        Assert.Empty(Upcoming(await Repo("!TD 오늘 오후 3시 보고서 제출", "!CD 내일 워크샵")));

    [Fact]
    public async Task RemindersComeBackSoonestFirst()
    {
        var repository = await Repo(
            "!CD 8월 12일 오후 3시 늦은 회의 30분 전 알림",
            "!TD 오늘 오후 3시 빠른 할일 30분 전 알림");

        var reminders = Upcoming(repository);

        Assert.Equal(2, reminders.Count);
        Assert.Equal("빠른 할일", reminders[0].Title);
        Assert.Equal("늦은 회의", reminders[1].Title);
    }

    /// <summary>Every round of a repeating entry is its own alarm, with its own moment.</summary>
    [Fact]
    public async Task ARepeatingEventGivesOneReminderPerOccurrence()
    {
        var repository = await Repo("!CD 매주 월요일 오전 10시 주간회의 30분 전 알림");

        var reminders = ReminderPlanner.Upcoming(repository, Now, TimeSpan.FromDays(21));

        Assert.Equal(3, reminders.Count);
        Assert.All(reminders, r => Assert.Equal(new TimeSpan(9, 30, 0), r.At.TimeOfDay));
        Assert.Equal(3, reminders.Select(r => r.AlarmId).Distinct().Count());
    }

    [Fact]
    public async Task NothingBeyondTheHorizonComesBack()
    {
        var repository = await Repo("!CD 8월 20일 오후 3시 먼 회의 30분 전 알림");

        Assert.Empty(Upcoming(repository, TimeSpan.FromDays(2)));
        Assert.Single(Upcoming(repository, TimeSpan.FromDays(30)));
    }

    [Fact]
    public async Task TheListIsBounded()
    {
        var repository = await Repo("!CD 매일 오전 10시 스탠드업 30분 전 알림");

        Assert.Equal(5, ReminderPlanner.Upcoming(repository, Now, TimeSpan.FromDays(60), max: 5).Count);
    }

    /// <summary>A reminder already gone off is not upcoming, however recently it passed.</summary>
    [Fact]
    public async Task WhatHasAlreadyPassedIsNotUpcoming()
    {
        var repository = await Repo("!TD 오늘 오전 9시 15분 지난 할일 30분 전 알림");

        Assert.Empty(Upcoming(repository));
    }

    [Fact]
    public async Task ButItIsOverdue()
    {
        var repository = await Repo("!TD 오늘 오전 9시 15분 지난 할일 30분 전 알림");

        var overdue = ReminderPlanner.Overdue(repository, Now, TimeSpan.FromHours(1));

        Assert.Equal("지난 할일", Assert.Single(overdue).Title);
    }

    /// <summary>
    /// Long past is not worth saying. A notification arriving hours late is noise, and the
    /// thing it was about has already happened.
    /// </summary>
    [Fact]
    public async Task LongPastIsLetGo()
    {
        var repository = await Repo("!TD 오늘 오전 9시 15분 지난 할일 30분 전 알림");

        Assert.Empty(ReminderPlanner.Overdue(repository, Now.AddHours(6), TimeSpan.FromMinutes(15)));
    }

    [Fact]
    public async Task ACompletedTodoStopsReminding()
    {
        var repository = await Repo("!TD 오늘 오후 3시 보고서 제출 30분 전 알림");
        await repository.ToggleTodoAsync(repository.Todos[0].Id, Now);

        Assert.Empty(Upcoming(repository));
    }

    /// <summary>The id has to survive recalculation, or every pass would stack a new alarm.</summary>
    [Fact]
    public async Task TheAlarmIdIsStableAcrossRecalculation()
    {
        var repository = await Repo("!CD 오늘 오후 3시 팀 회의 30분 전 알림");

        Assert.Equal(Upcoming(repository)[0].AlarmId, Upcoming(repository)[0].AlarmId);
    }

    [Fact]
    public void AnEmptyWorkspaceIsQuiet()
    {
        var repository = new WorkspaceRepository(new InMemoryStore());

        Assert.Empty(Upcoming(repository));
        Assert.Empty(ReminderPlanner.Overdue(repository, Now, TimeSpan.FromHours(1)));
    }

    [Fact]
    public async Task TheDescriptionSaysWhenTheThingItselfIs()
    {
        var repository = await Repo("!CD 오늘 오후 3시 팀 회의 30분 전 알림");

        Assert.Contains("8월 10일", ReminderPlanner.Describe(Upcoming(repository)[0]));
    }
}

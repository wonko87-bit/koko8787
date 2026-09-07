using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// A week read back as report material. What lands in which section, and that nothing is
/// invented to fill one.
/// </summary>
public class WeeklyReportTests
{
    // Monday 2026-08-17 to Sunday 2026-08-23.
    private static readonly DateTime Mon = new(2026, 8, 17);
    private static readonly DateTime Sun = new(2026, 8, 23);

    /// <summary>
    /// Loaded straight in, not imported: an import strips the origin off a copied-in
    /// meeting, which is right for a file from another machine and wrong for a test of one.
    /// </summary>
    private static async Task<WorkspaceRepository> Repository(params object[] entries)
    {
        var workspace = new Flowdeck.Core.Storage.Workspace();
        foreach (var entry in entries)
        {
            switch (entry)
            {
                case TodoItem todo: workspace.Todos.Add(todo); break;
                case CalendarEvent calendarEvent: workspace.Events.Add(calendarEvent); break;
            }
        }

        var repository = new WorkspaceRepository(new InMemoryStore());
        await repository.ReplaceAsync(workspace);
        return repository;
    }

    private static TodoItem Done(string title, DateTime at, string notes = "") => new()
    {
        Title = title,
        IsDone = true,
        CompletedAt = at,
        CreatedAt = at.AddDays(-3),
        UpdatedAt = at,
        Notes = notes,
    };

    private static TodoItem Open(string title, DateTime? due, string notes = "", bool hasTime = false) => new()
    {
        Title = title,
        DueAt = due,
        HasTime = hasTime,
        CreatedAt = Mon.AddDays(-30),
        UpdatedAt = Mon.AddDays(-30),
        Notes = notes,
    };

    private static CalendarEvent Meeting(string title, DateTime start, string notes, bool adopted = true) => new()
    {
        Title = title,
        Start = start,
        End = start.AddHours(1),
        Notes = notes,
        Origin = adopted ? new ExternalOrigin("M-" + title, start) : null,
    };

    // ---- the week ----------------------------------------------------------

    [Theory]
    [InlineData(2026, 8, 17, DayOfWeek.Monday, 2026, 8, 17)]   // a Monday, weeks on Monday
    [InlineData(2026, 8, 20, DayOfWeek.Monday, 2026, 8, 17)]   // a Thursday
    [InlineData(2026, 8, 23, DayOfWeek.Monday, 2026, 8, 17)]   // the Sunday still belongs to it
    [InlineData(2026, 8, 23, DayOfWeek.Sunday, 2026, 8, 23)]   // weeks on Sunday: a new one
    [InlineData(2026, 8, 20, DayOfWeek.Sunday, 2026, 8, 16)]
    public void TheWeekFollowsTheFirstDaySetting(int y, int m, int d, DayOfWeek first, int fy, int fm, int fd)
    {
        var (from, to) = WeeklyReport.WeekOf(new DateTime(y, m, d, 15, 30, 0), first);

        Assert.Equal(new DateTime(fy, fm, fd), from);
        Assert.Equal(from.AddDays(6), to);
    }

    // ---- sections ------------------------------------------------------------

    [Fact]
    public async Task WhatWasTickedOffThisWeekIsCompleted()
    {
        var repository = await Repository(
            Done("이번 주에 한 것", Mon.AddDays(2)),
            Done("지난주에 한 것", Mon.AddDays(-2)),
            Done("다음 주에 할 것", Sun.AddDays(2)),
            Open("아직 안 한 것", Mon.AddDays(3)));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Equal("이번 주에 한 것", Assert.Single(report.Completed).Title);
    }

    [Fact]
    public async Task ASundayCompletionStillCountsAsThisWeek()
    {
        var repository = await Repository(Done("일요일 저녁", Sun.AddHours(23).AddMinutes(30)));

        Assert.Single(WeeklyReport.Build(repository, Mon, Sun).Completed);
    }

    [Fact]
    public async Task AMeetingWithANoteIsReportedAndItsLinesArePulledOut()
    {
        var repository = await Repository(
            Meeting("설계 리뷰", Mon.AddDays(2).AddHours(14),
                "손실 해석 결과 공유\n결정: 코어 재질 B안으로 확정\n할일: 재해석, 다음주 수요일까지\n이슈: 열 해석 미착수\n다음: 9/2 후속 리뷰"),
            Meeting("주간 회의", Mon.AddHours(10), "결정: 9월 일정 1주 연기"));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Equal(new[] { "주간 회의", "설계 리뷰" }, report.Meetings.Select(m => m.Title));
        Assert.Equal(
            new[] { ("주간 회의", "9월 일정 1주 연기"), ("설계 리뷰", "코어 재질 B안으로 확정") },
            report.Decisions.Select(d => (d.Source, d.Text)));
        Assert.Equal(("설계 리뷰", "재해석, 다음주 수요일까지"), (Assert.Single(report.Actions).Source, report.Actions[0].Text));
        Assert.Equal("열 해석 미착수", Assert.Single(report.Issues).Text);
        Assert.Equal("9/2 후속 리뷰", Assert.Single(report.NextSteps).Text);
    }

    [Fact]
    public async Task AMeetingNobodyWroteAnythingOnIsNotReported()
    {
        // Taken in, but the note was never written. There is nothing to say about it.
        var repository = await Repository(Meeting("빈 회의", Mon.AddHours(10), ""));

        Assert.Empty(WeeklyReport.Build(repository, Mon, Sun).Meetings);
    }

    [Fact]
    public async Task AnOrdinaryEventWrittenUpLikeAMeetingCountsAsOne()
    {
        var repository = await Repository(
            Meeting("협력사 미팅", Mon.AddDays(1).AddHours(15), "결정: 납기 2주 연장", adopted: false));

        Assert.Equal("협력사 미팅", Assert.Single(WeeklyReport.Build(repository, Mon, Sun).Meetings).Title);
    }

    [Fact]
    public async Task AnOrdinaryEventWithAPlainNoteIsNotAMeeting()
    {
        var repository = await Repository(
            Meeting("치과", Mon.AddDays(1).AddHours(15), "주차는 건물 뒤", adopted: false));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Empty(report.Meetings);
        Assert.Empty(report.Runs);
    }

    [Fact]
    public async Task ARunIsReadOffItsConditionsResultsAndConclusion()
    {
        var repository = await Repository(
            Done("모터A 손실 해석", Mon.AddDays(4),
                "조건: 코어 B안, 120A, 3000rpm\n결과: 철손 42W\n결론: A안 대비 8% 감소\n파일: D:\\해석\\run_0821.aedt\n할일: 열 해석 추가"));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        var run = Assert.Single(report.Runs);
        Assert.Equal("모터A 손실 해석", run.Title);
        Assert.Equal(new[] { "코어 B안, 120A, 3000rpm" }, run.Outline.Conditions);
        Assert.Equal(new[] { "A안 대비 8% 감소" }, run.Outline.Conclusions);

        // Its action item joins the week's list, and its file the week's files.
        Assert.Equal(("모터A 손실 해석", "열 해석 추가"), (Assert.Single(report.Actions).Source, report.Actions[0].Text));
        Assert.Equal(new[] { @"D:\해석\run_0821.aedt" }, report.Files);

        // Being done, it is also in the completed list — that is not double counting, it
        // is two different things being true of it.
        Assert.Single(report.Completed);
    }

    [Fact]
    public async Task APlainTodoWithTheOpenersFeedsTheListsAcrossTheWeek()
    {
        // A phone call is a todo, not a meeting. What was decided on it is still a decision.
        var repository = await Repository(
            Done("협력사 통화", Mon.AddDays(1), "납기 논의\n결정: 납기 2주 연장\n할일: 계약서 수정\n이슈: 단가 미합의\n다음: 8/28 재통화"),
            Open("이번 주 기한인데 아직", Mon.AddDays(4), "이슈: 자료 미수령"),
            Open("한참 뒤 기한", Sun.AddDays(30), "결정: 이건 이번 주 것이 아님"));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        // No block of its own: it is neither a meeting nor a run.
        Assert.Empty(report.Meetings);
        Assert.Empty(report.Runs);

        Assert.Equal(("협력사 통화", "납기 2주 연장"), (Assert.Single(report.Decisions).Source, report.Decisions[0].Text));
        Assert.Equal("계약서 수정", Assert.Single(report.Actions).Text);
        Assert.Equal(new[] { "단가 미합의", "자료 미수령" }, report.Issues.Select(i => i.Text));
        Assert.Equal("8/28 재통화", Assert.Single(report.NextSteps).Text);
    }

    [Fact]
    public async Task MeetingsComeBeforeTodosInTheListsAcrossTheWeek()
    {
        var repository = await Repository(
            Done("월요일 통화", Mon, "결정: 통화에서 정함"),
            Meeting("금요일 회의", Mon.AddDays(4).AddHours(10), "결정: 회의에서 정함"));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Equal(new[] { "금요일 회의", "월요일 통화" }, report.Decisions.Select(d => d.Source));
    }

    [Fact]
    public async Task AnAdoptedMeetingIsAMeetingEvenIfItsNoteReadsLikeARun()
    {
        var repository = await Repository(
            Meeting("해석 결과 리뷰", Mon.AddHours(10), "조건: 120A\n결과: 42W\n결정: 채택"));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Single(report.Meetings);
        Assert.Empty(report.Runs);
    }

    [Fact]
    public async Task FilesHandedOverThisWeekAreListedOnce()
    {
        var arrived = Open("[읽기] 보고서", Mon.AddDays(10), "파일: C:\\관리함\\보고서.pdf");
        arrived.CreatedAt = Mon.AddDays(1);

        var repository = await Repository(
            arrived,
            Done("같은 파일 또", Mon.AddDays(2), "파일: c:\\관리함\\보고서.PDF"),
            Open("오래된 것", null, "파일: C:\\관리함\\옛날.pdf"));

        Assert.Equal(new[] { @"C:\관리함\보고서.pdf" }, WeeklyReport.Build(repository, Mon, Sun).Files);
    }

    [Fact]
    public async Task WhatIsOverdueAndWhatIsNextAreKeptApart()
    {
        var repository = await Repository(
            Open("이번 주였는데", Mon.AddDays(3)),
            Open("한참 전이었는데", Mon.AddDays(-20)),
            Open("다음 주", Sun.AddDays(3)),
            Open("한 달 뒤", Sun.AddDays(20)),
            Open("언젠가", null));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Equal(new[] { "한참 전이었는데", "이번 주였는데" }, report.Unfinished.Select(t => t.Title));
        Assert.Equal(new[] { "다음 주" }, report.Upcoming.Select(t => t.Title));
    }

    [Fact]
    public async Task AReversedRangeIsPutTheRightWayRound()
    {
        var repository = await Repository(Done("한 것", Mon.AddDays(2)));

        var report = WeeklyReport.Build(repository, Sun, Mon);

        Assert.Equal(Mon, report.From);
        Assert.Equal(Sun, report.To);
        Assert.Single(report.Completed);
    }

    [Fact]
    public async Task WhatWasOwedToAMeetingIsListedUnderItDoneOrNot()
    {
        var meeting = Meeting("설계 리뷰", Mon.AddDays(2).AddHours(14), "손실 공유\n결정: B안 확정");
        var repository = await Repository(meeting);
        var slides = await repository.AttachTodoAsync(meeting.Id, "발표 자료 작성", Mon);
        await repository.AttachTodoAsync(meeting.Id, "회의실 예약", Mon);
        await repository.ToggleTodoAsync(slides.Id, Mon.AddDays(1));

        var report = WeeklyReport.Build(repository, Mon, Sun);

        var noted = Assert.Single(report.Meetings);
        Assert.Equal(new[] { ("회의실 예약", false), ("발표 자료 작성", true) }, noted.AttachedTodos.Select(t => (t.Title, t.Done)));

        var text = report.Render().Replace("\r\n", "\n");
        Assert.Contains("    결정: B안 확정\n    준비: 회의실 예약 (미완)\n    준비: 발표 자료 작성 (완료)", text);
    }

    [Fact]
    public async Task AMeetingWithNothingWrittenButSomethingAttachedIsStillReported()
    {
        // Nothing was said about it, but something was done for it.
        var meeting = Meeting("빈 회의", Mon.AddHours(10), "");
        var repository = await Repository(meeting);
        await repository.AttachTodoAsync(meeting.Id, "자료", Mon);

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.Equal("빈 회의", Assert.Single(report.Meetings).Title);
    }

    [Fact]
    public async Task AnOrdinaryEventWithSomethingAttachedCountsAsAMeeting()
    {
        var outing = Meeting("협력사 방문", Mon.AddDays(3).AddHours(10), "", adopted: false);
        var repository = await Repository(outing);
        await repository.AttachTodoAsync(outing.Id, "샘플 챙기기", Mon);

        Assert.Single(WeeklyReport.Build(repository, Mon, Sun).Meetings);
    }

    // ---- text ----------------------------------------------------------------

    [Fact]
    public async Task AnEmptyWeekSaysSo()
    {
        var repository = await Repository();

        var report = WeeklyReport.Build(repository, Mon, Sun);

        Assert.True(report.IsEmpty);
        Assert.Contains("이 기간에 기록된 항목이 없습니다", report.Render());
    }

    [Fact]
    public async Task TheTextHasOnlyTheSectionsWithSomethingInThem()
    {
        var repository = await Repository(
            Done("리포트 검토", Mon.AddDays(3)),
            Meeting("설계 리뷰", Mon.AddDays(2).AddHours(14), "손실 공유\n결정: B안 확정"),
            Open("재해석", Sun.AddDays(3), hasTime: false));

        var text = WeeklyReport.Build(repository, Mon, Sun).Render();

        Assert.StartsWith("주간보고 초안 · 2026-08-17 (월) ~ 2026-08-23 (일)", text);
        Assert.Contains("■ 완료 (1)\n  - 리포트 검토 · 8/20 (목)", text.Replace("\r\n", "\n"));
        Assert.Contains("■ 회의 (1)\n  설계 리뷰 · 8/19 (수) 14:00\n    손실 공유\n    결정: B안 확정", text.Replace("\r\n", "\n"));
        Assert.Contains("■ 이번 주 결정 (1)\n  - [설계 리뷰] B안 확정", text.Replace("\r\n", "\n"));
        Assert.Contains("■ 다음 주 (1)\n  - 재해석 · 8/26 (수)", text.Replace("\r\n", "\n"));

        Assert.DoesNotContain("■ 이슈", text);
        Assert.DoesNotContain("■ 해석", text);
        Assert.DoesNotContain("■ 미완료", text);
    }

    [Fact]
    public async Task AFileIsNamedByItsNameNotItsPath()
    {
        var repository = await Repository(Done("읽기", Mon.AddDays(1), "파일: C:\\아주\\깊은\\폴더\\보고서.pdf"));

        var text = WeeklyReport.Build(repository, Mon, Sun).Render();

        Assert.Contains("보고서.pdf", text);
        Assert.DoesNotContain(@"C:\아주", text);
    }
}

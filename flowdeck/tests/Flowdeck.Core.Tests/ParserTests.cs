using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Xunit;

namespace Flowdeck.Core.Tests;

public class ParserTests
{
    // Thursday, 6 August 2026, 10:00. Every relative assertion below hangs off this.
    private static readonly DateTime Now = new(2026, 8, 6, 10, 0, 0);

    private static ParsedEntry Parse(string input) => new NaturalLanguageParser().Parse(input, Now);

    [Fact]
    public void EmptyInput_ProducesEmptyEntry()
    {
        var entry = Parse("   ");

        Assert.True(entry.IsEmpty);
        Assert.False(entry.HasDate);
    }

    [Fact]
    public void PlainText_HasNoScheduleAndKeepsWholeTitle()
    {
        var entry = Parse("영수증 정리");

        Assert.Equal("영수증 정리", entry.Title);
        Assert.False(entry.HasDate);
        Assert.Null(entry.Start);
    }

    // ---- dates -------------------------------------------------------------

    [Theory]
    [InlineData("오늘 정리", 2026, 8, 6)]
    [InlineData("내일 정리", 2026, 8, 7)]
    [InlineData("모레 정리", 2026, 8, 8)]
    [InlineData("글피 정리", 2026, 8, 9)]
    [InlineData("어제 정리", 2026, 8, 5)]
    public void RelativeDayWords(string input, int y, int m, int d) =>
        Assert.Equal(new DateTime(y, m, d), Parse(input).Start);

    [Theory]
    [InlineData("8월 12일 워크샵", 2026, 8, 12)]
    [InlineData("2026-12-25 크리스마스", 2026, 12, 25)]
    [InlineData("9/15 점검", 2026, 9, 15)]
    [InlineData("3일 후 제출", 2026, 8, 9)]
    [InlineData("2주 후 검토", 2026, 8, 20)]
    public void AbsoluteAndOffsetDates(string input, int y, int m, int d) =>
        Assert.Equal(new DateTime(y, m, d), Parse(input).Start);

    [Fact]
    public void MonthDayInThePast_RollsToNextYear() =>
        Assert.Equal(new DateTime(2027, 3, 2), Parse("3월 2일 계약 갱신").Start);

    [Fact]
    public void BareDayNumber_UsesThisMonthWhenStillAhead() =>
        Assert.Equal(new DateTime(2026, 8, 20), Parse("20일 정산").Start);

    [Fact]
    public void BareDayNumber_RollsToNextMonthWhenAlreadyPast() =>
        Assert.Equal(new DateTime(2026, 9, 2), Parse("2일 정산").Start);

    [Theory]
    [InlineData("금요일 회고", 2026, 8, 7)]      // next Friday, today is Thursday
    [InlineData("목요일 회고", 2026, 8, 6)]      // today counts
    [InlineData("이번주 금요일 회고", 2026, 8, 7)]
    [InlineData("다음주 월요일 회고", 2026, 8, 10)]
    [InlineData("다다음주 월요일 회고", 2026, 8, 17)]
    [InlineData("주말 등산", 2026, 8, 8)]
    public void Weekdays(string input, int y, int m, int d) =>
        Assert.Equal(new DateTime(y, m, d), Parse(input).Start);

    [Fact]
    public void SundayFirstWeeksShiftWhatThisWeekMeans()
    {
        // Now is Thursday 6 Aug. Monday-first, this week opened on the 3rd, so
        // "이번주 일요일" is the 9th. Sunday-first, it opened on the 2nd — already past.
        var mondayFirst = new NaturalLanguageParser { FirstDayOfWeek = DayOfWeek.Monday };
        var sundayFirst = new NaturalLanguageParser { FirstDayOfWeek = DayOfWeek.Sunday };

        Assert.Equal(new DateTime(2026, 8, 9), mondayFirst.Parse("이번주 일요일 정리", Now).Start);
        Assert.Equal(new DateTime(2026, 8, 2), sundayFirst.Parse("이번주 일요일 정리", Now).Start);

        Assert.Equal(new DateTime(2026, 8, 16), mondayFirst.Parse("다음주 일요일 정리", Now).Start);
        Assert.Equal(new DateTime(2026, 8, 9), sundayFirst.Parse("다음주 일요일 정리", Now).Start);
    }

    [Fact]
    public void WeekStartDoesNotDisturbMidWeekDays()
    {
        var sundayFirst = new NaturalLanguageParser { FirstDayOfWeek = DayOfWeek.Sunday };

        // Friday sits in the same week either way, as does a bare weekday.
        Assert.Equal(new DateTime(2026, 8, 7), sundayFirst.Parse("이번주 금요일 회고", Now).Start);
        Assert.Equal(new DateTime(2026, 8, 7), sundayFirst.Parse("금요일 회고", Now).Start);
        Assert.Equal(new DateTime(2026, 8, 8), sundayFirst.Parse("주말 등산", Now).Start);
    }

    [Fact]
    public void NextMonthWithDay() =>
        Assert.Equal(new DateTime(2026, 9, 5), Parse("다음달 5일 정기점검").Start);

    // ---- times -------------------------------------------------------------

    [Fact]
    public void MeridiemTime()
    {
        var entry = Parse("내일 오후 9시 운동");

        Assert.Equal(new DateTime(2026, 8, 7, 21, 0, 0), entry.Start);
        Assert.True(entry.HasTime);
        Assert.False(entry.IsAllDay);
    }

    [Theory]
    [InlineData("내일 오전 9시 30분 통화", 9, 30)]
    [InlineData("내일 오후 3시 반 통화", 15, 30)]
    [InlineData("내일 14:30 통화", 14, 30)]
    [InlineData("내일 저녁 7시 통화", 19, 0)]
    [InlineData("내일 새벽 5시 통화", 5, 0)]
    [InlineData("내일 밤 11시 통화", 23, 0)]
    [InlineData("내일 9:15am 통화", 9, 15)]
    [InlineData("내일 8pm 통화", 20, 0)]
    public void TimeFormats(string input, int hour, int minute) =>
        Assert.Equal(new DateTime(2026, 8, 7, hour, minute, 0), Parse(input).Start);

    [Theory]
    [InlineData("내일 점심 약속", 12)]
    [InlineData("내일 저녁 약속", 18)]
    [InlineData("내일 정오 약속", 12)]
    public void StandaloneDayParts(string input, int hour) =>
        Assert.Equal(new DateTime(2026, 8, 7, hour, 0, 0), Parse(input).Start);

    [Fact]
    public void BareEveningHour_IsReadAsAfternoon() =>
        Assert.Equal(new DateTime(2026, 8, 7, 15, 0, 0), Parse("내일 3시 회의").Start);

    [Fact]
    public void BareMorningHour_StaysInTheMorning() =>
        Assert.Equal(new DateTime(2026, 8, 7, 9, 0, 0), Parse("내일 9시 회의").Start);

    [Fact]
    public void ExplicitMeridiem_IsNeverShifted() =>
        Assert.Equal(new DateTime(2026, 8, 7, 3, 0, 0), Parse("내일 오전 3시 배포").Start);

    [Fact]
    public void LiteralHoursWhenTheAssumptionIsOff()
    {
        var parser = new NaturalLanguageParser { AssumeAfternoonForBareHours = false };

        Assert.Equal(new DateTime(2026, 8, 7, 3, 0, 0), parser.Parse("내일 3시 회의", Now).Start);
    }

    [Fact]
    public void TimeWithoutDate_UsesTodayWhenStillAhead() =>
        Assert.Equal(new DateTime(2026, 8, 6, 14, 0, 0), Parse("오후 2시 통화").Start);

    [Fact]
    public void TimeWithoutDate_RollsToTomorrowWhenAlreadyPast() =>
        Assert.Equal(new DateTime(2026, 8, 7, 8, 0, 0), Parse("오전 8시 조깅").Start);

    // ---- ranges and durations ----------------------------------------------

    [Fact]
    public void TildeRange()
    {
        var entry = Parse("내일 오후 3시~5시 회의");

        Assert.Equal(new DateTime(2026, 8, 7, 15, 0, 0), entry.Start);
        Assert.Equal(new DateTime(2026, 8, 7, 17, 0, 0), entry.End);
    }

    [Fact]
    public void ParticleRange()
    {
        var entry = Parse("내일 오전 10시부터 오후 1시까지 워크샵");

        Assert.Equal(new DateTime(2026, 8, 7, 10, 0, 0), entry.Start);
        Assert.Equal(new DateTime(2026, 8, 7, 13, 0, 0), entry.End);
    }

    [Fact]
    public void BareEndHour_SharesTheStartsHalfOfTheDay()
    {
        var entry = Parse("내일 오전 10시~1시 회의");

        Assert.Equal(new DateTime(2026, 8, 7, 10, 0, 0), entry.Start);
        Assert.Equal(new DateTime(2026, 8, 7, 13, 0, 0), entry.End);
    }

    [Fact]
    public void MultiDayAllDayRange()
    {
        var entry = Parse("8월 10일 ~ 8월 12일 워크샵");

        Assert.Equal(new DateTime(2026, 8, 10), entry.Start);
        Assert.Equal(new DateTime(2026, 8, 12), entry.End);
        Assert.True(entry.IsAllDay);
        Assert.Equal("워크샵", entry.Title);
    }

    [Fact]
    public void Duration_SetsTheEnd()
    {
        var entry = Parse("내일 오후 2시 2시간 워크샵");

        Assert.Equal(new DateTime(2026, 8, 7, 14, 0, 0), entry.Start);
        Assert.Equal(new DateTime(2026, 8, 7, 16, 0, 0), entry.End);
        Assert.Equal("워크샵", entry.Title);
    }

    [Fact]
    public void MinuteDuration()
    {
        var entry = Parse("내일 오후 2시 30분간 통화");

        Assert.Equal(new DateTime(2026, 8, 7, 14, 30, 0), entry.End);
    }

    [Fact]
    public void TimedEntryWithoutEnd_DefaultsToOneHour()
    {
        var entry = Parse("내일 오후 3시 회의");

        Assert.Equal(new DateTime(2026, 8, 7, 16, 0, 0), entry.End);
    }

    // ---- recurrence, priority, tags, reminders -----------------------------

    [Fact]
    public void WeeklyRecurrenceTakesItsDayFromTheStart()
    {
        var entry = Parse("매주 화요일 오전 10시 주간회의");

        Assert.Equal(RecurrenceKind.Weekly, entry.Recurrence.Kind);
        Assert.Equal(1, entry.Recurrence.Interval);
        Assert.Equal(new[] { DayOfWeek.Tuesday }, entry.Recurrence.DaysOfWeek);
        Assert.Equal(new DateTime(2026, 8, 11, 10, 0, 0), entry.Start);
        Assert.Equal("주간회의", entry.Title);
    }

    [Theory]
    [InlineData("매일 오전 7시 스트레칭", RecurrenceKind.Daily, 1)]
    [InlineData("격주 목요일 회식", RecurrenceKind.Weekly, 2)]
    [InlineData("매월 25일 월세 이체", RecurrenceKind.Monthly, 1)]
    [InlineData("매년 3월 2일 면허 갱신", RecurrenceKind.Yearly, 1)]
    [InlineData("평일 오전 9시 스탠드업", RecurrenceKind.Weekdays, 1)]
    [InlineData("3일마다 화분 물주기", RecurrenceKind.Daily, 3)]
    public void RecurrenceWords(string input, RecurrenceKind kind, int interval)
    {
        var entry = Parse(input);

        Assert.Equal(kind, entry.Recurrence.Kind);
        Assert.Equal(interval, entry.Recurrence.Interval);
    }

    [Theory]
    [InlineData("!1 보고서 제출", Priority.Urgent)]
    [InlineData("!2 보고서 제출", Priority.High)]
    [InlineData("!4 보고서 제출", Priority.Low)]
    public void PrioritySymbols(string input, Priority expected)
    {
        var entry = Parse(input);

        Assert.Equal(expected, entry.Priority);
        Assert.Equal("보고서 제출", entry.Title);
    }

    [Fact]
    public void PriorityWords_SetPriorityButStayInTheTitle()
    {
        var entry = Parse("긴급 서버 점검");

        Assert.Equal(Priority.Urgent, entry.Priority);
        Assert.Equal("긴급 서버 점검", entry.Title);
    }

    [Fact]
    public void Tags()
    {
        var entry = Parse("내일 오후 3시 스프린트 리뷰 #업무 #팀");

        Assert.Equal(new[] { "업무", "팀" }, entry.Tags);
        Assert.Equal("스프린트 리뷰", entry.Title);
    }

    [Theory]
    [InlineData("내일 오후 3시 회의 30분 전 알림", 30)]
    [InlineData("내일 오후 3시 회의 1시간 전 알림", 60)]
    public void Reminders(string input, int minutes) =>
        Assert.Equal(minutes, Parse(input).ReminderMinutesBefore);

    // ---- title cleanup -----------------------------------------------------

    [Theory]
    [InlineData("내일 오후 3시에 팀 회의", "팀 회의")]
    [InlineData("내일까지 보고서 제출", "보고서 제출")]
    [InlineData("8월 12일 오전 10시 고객사 미팅", "고객사 미팅")]
    [InlineData("내일 오후 3시", "(제목 없음)")]
    public void TitleKeepsOnlyWhatIsLeft(string input, string expected) =>
        Assert.Equal(expected, Parse(input).Title);

    [Fact]
    public void OrdinaryWordsContainingDigitsAreNotEatenAsDates()
    {
        var entry = Parse("경리팀 미팅 준비");

        Assert.Equal("경리팀 미팅 준비", entry.Title);
        Assert.False(entry.HasDate);
    }

    [Fact]
    public void HoursWordIsNotMistakenForAClockTime()
    {
        var entry = Parse("2시간 집중 작업");

        Assert.False(entry.HasTime);
        Assert.Equal("집중 작업", entry.Title);
    }
}

using Flowdeck.Core.Models;
using Flowdeck.Core.Services;
using Xunit;

namespace Flowdeck.Core.Tests;

public class RecurrenceTests
{
    private static List<DateTime> Expand(DateTime start, Recurrence rule, DateTime from, DateTime to) =>
        RecurrenceExpander.Occurrences(start, rule, from, to).ToList();

    [Fact]
    public void NoRule_YieldsTheSingleDateWhenItIsInWindow()
    {
        var start = new DateTime(2026, 8, 11, 10, 0, 0);

        Assert.Equal(new[] { start }, Expand(start, Recurrence.None, new DateTime(2026, 8, 1), new DateTime(2026, 8, 31)));
        Assert.Empty(Expand(start, Recurrence.None, new DateTime(2026, 9, 1), new DateTime(2026, 9, 30)));
    }

    [Fact]
    public void WeeklyOnATuesday()
    {
        var start = new DateTime(2026, 8, 11, 10, 0, 0);
        var rule = new Recurrence { Kind = RecurrenceKind.Weekly, Interval = 1, DaysOfWeek = { DayOfWeek.Tuesday } };

        var dates = Expand(start, rule, new DateTime(2026, 8, 1), new DateTime(2026, 8, 31));

        Assert.Equal(
            new[]
            {
                new DateTime(2026, 8, 11, 10, 0, 0),
                new DateTime(2026, 8, 18, 10, 0, 0),
                new DateTime(2026, 8, 25, 10, 0, 0),
            },
            dates);
    }

    [Fact]
    public void EveryOtherWeekSkipsTheWeekBetween()
    {
        var start = new DateTime(2026, 8, 6, 19, 0, 0); // Thursday
        var rule = new Recurrence { Kind = RecurrenceKind.Weekly, Interval = 2 };

        var dates = Expand(start, rule, new DateTime(2026, 8, 1), new DateTime(2026, 8, 31));

        Assert.Equal(
            new[] { new DateTime(2026, 8, 6, 19, 0, 0), new DateTime(2026, 8, 20, 19, 0, 0) },
            dates);
    }

    [Fact]
    public void EveryThirdDay()
    {
        var start = new DateTime(2026, 8, 6);
        var rule = new Recurrence { Kind = RecurrenceKind.Daily, Interval = 3 };

        Assert.Equal(
            new[]
            {
                new DateTime(2026, 8, 6),
                new DateTime(2026, 8, 9),
                new DateTime(2026, 8, 12),
                new DateTime(2026, 8, 15),
            },
            Expand(start, rule, new DateTime(2026, 8, 6), new DateTime(2026, 8, 15)));
    }

    [Fact]
    public void WeekdaysSkipTheWeekend()
    {
        var start = new DateTime(2026, 8, 6, 9, 0, 0); // Thursday
        var rule = new Recurrence { Kind = RecurrenceKind.Weekdays };

        var dates = Expand(start, rule, new DateTime(2026, 8, 6), new DateTime(2026, 8, 12));

        Assert.Equal(
            new[]
            {
                new DateTime(2026, 8, 6, 9, 0, 0),
                new DateTime(2026, 8, 7, 9, 0, 0),
                new DateTime(2026, 8, 10, 9, 0, 0),
                new DateTime(2026, 8, 11, 9, 0, 0),
                new DateTime(2026, 8, 12, 9, 0, 0),
            },
            dates);
    }

    [Fact]
    public void MonthlyClampsToTheShortestMonth()
    {
        var start = new DateTime(2026, 1, 31);
        var rule = new Recurrence { Kind = RecurrenceKind.Monthly };

        var dates = Expand(start, rule, new DateTime(2026, 1, 1), new DateTime(2026, 4, 30));

        Assert.Equal(
            new[]
            {
                new DateTime(2026, 1, 31),
                new DateTime(2026, 2, 28),
                new DateTime(2026, 3, 31),
                new DateTime(2026, 4, 30),
            },
            dates);
    }

    [Fact]
    public void UntilStopsTheSeries()
    {
        var start = new DateTime(2026, 8, 6);
        var rule = new Recurrence
        {
            Kind = RecurrenceKind.Daily,
            Until = new DateTime(2026, 8, 8),
        };

        Assert.Equal(3, Expand(start, rule, new DateTime(2026, 8, 1), new DateTime(2026, 8, 31)).Count);
    }

    [Fact]
    public void NextFindsTheFollowingOccurrence()
    {
        var start = new DateTime(2026, 8, 6, 7, 0, 0);
        var rule = new Recurrence { Kind = RecurrenceKind.Daily };

        Assert.Equal(new DateTime(2026, 8, 7, 7, 0, 0), RecurrenceExpander.Next(start, rule, start));
    }

    [Fact]
    public void NextReturnsNothingForAOneOff() =>
        Assert.Null(RecurrenceExpander.Next(new DateTime(2026, 8, 6), Recurrence.None, new DateTime(2026, 8, 6)));
}

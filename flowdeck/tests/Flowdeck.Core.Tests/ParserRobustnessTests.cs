using Flowdeck.Core.Parsing;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// The parser runs on every keystroke to drive the live preview, so it meets every
/// half-finished prefix of whatever the user is typing — not just complete sentences.
/// These walk each sample one character at a time and demand that nothing throws.
/// </summary>
public class ParserRobustnessTests
{
    private static readonly DateTime Now = new(2026, 8, 7, 14, 30, 0);

    public static TheoryData<string> Samples() => new()
    {
        "내일 오후 3시 팀 회의",
        "!TD 금요일까지 보고서 제출 #업무 #보고",
        "매주 월,목 3분할 운동 - 상체",
        "매주 화요일 오전 10시 주간회의 #팀 30분 전 알림",
        "8월 10일 ~ 8월 12일 워크샵 !CD",
        "!1 긴급 서버 점검 2시간",
        "다음주 월요일 오후 2시부터 5시까지 워크샵",
        "매달 25일 월세 이체",
        "2029-01-15 장기 계약 갱신",
        "2월 29일 윤년",
        "12시간 30분 전 알림 99일 후",
        "9/15 점검 #a #b #c",
        "평일 오전 9시 스탠드업",
        "31일 정산",
        "24시 이상한 시간",
        "0시 0분",
        "~ 5시",
        "!!!! ### @@@ /// ~~~",
        "999일 후 9999년 99월 99일",
    };

    [Theory]
    [MemberData(nameof(Samples))]
    public void EveryPrefixParsesWithoutThrowing(string input)
    {
        var parser = new NaturalLanguageParser();

        for (var length = 0; length <= input.Length; length++)
        {
            var prefix = input.Substring(0, length);
            var entry = parser.Parse(prefix, Now);

            // The preview reads all of these while the user is still typing.
            _ = entry.DescribeSchedule();
            _ = entry.DescribeTarget();
            _ = entry.Recurrence.Describe();
            _ = entry.Recurrence.DescribeShort(entry.Start);

            Assert.NotNull(entry.Title);
        }
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public void ResultsStayInsideTheCalendarsRange(string input)
    {
        var parser = new NaturalLanguageParser();

        for (var length = 0; length <= input.Length; length++)
        {
            var entry = parser.Parse(input.Substring(0, length), Now);
            if (!entry.Start.HasValue) continue;

            // The widget builds a month grid around whatever comes back, and paging off
            // either end of DateTime would throw where there is no caller to catch it.
            Assert.InRange(entry.Start.Value, Now.AddYears(-50), Now.AddYears(50));
            if (entry.End.HasValue) Assert.InRange(entry.End.Value, entry.Start.Value, Now.AddYears(50));
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\t\n")]
    public void BlankInputIsEmpty(string input)
    {
        var entry = new NaturalLanguageParser().Parse(input, Now);

        Assert.True(entry.IsEmpty);
        Assert.False(entry.HasDate);
    }
}

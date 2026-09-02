using Flowdeck.Core.Parsing;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// The line-opener convention: eight words that, at the start of a line and followed by a
/// colon, let a week of notes be read back as decisions, actions and conclusions.
/// </summary>
public class NoteOutlineTests
{
    private const string Meeting =
        "손실 해석 결과 공유. 3번 조건이 기준 초과라 재질 변경 논의\n" +
        "결정: 코어 재질 B안으로 확정\n" +
        "할일: 재질 변경 후 재해석, 다음주 수요일까지\n" +
        "이슈: 열 해석은 아직 미착수\n" +
        "다음: 9/2 후속 리뷰";

    private const string Run =
        "조건: 코어 B안, 120A, 3000rpm\n" +
        "결과: 철손 42W, 동손 118W, 효율 93.1%\n" +
        "결론: A안 대비 손실 8% 감소\n" +
        "파일: D:\\해석\\모터A\\run_0821.aedt\n" +
        "할일: 열 해석 추가, 이번주 금요일까지";

    // ---- the two shapes the convention was drawn for -----------------------

    [Fact]
    public void AMeetingNoteComesApartIntoItsParts()
    {
        var outline = NoteOutline.Parse(Meeting);

        Assert.Equal("손실 해석 결과 공유. 3번 조건이 기준 초과라 재질 변경 논의", outline.Summary);
        Assert.Equal(new[] { "코어 재질 B안으로 확정" }, outline.Decisions);
        Assert.Equal(new[] { "재질 변경 후 재해석, 다음주 수요일까지" }, outline.Actions);
        Assert.Equal(new[] { "열 해석은 아직 미착수" }, outline.Issues);
        Assert.Equal(new[] { "9/2 후속 리뷰" }, outline.Next);
        Assert.True(outline.HasStructure);
        Assert.False(outline.IsAnalysis);
    }

    [Fact]
    public void ARunNoteComesApartIntoItsParts()
    {
        var outline = NoteOutline.Parse(Run);

        Assert.Equal(new[] { "코어 B안, 120A, 3000rpm" }, outline.Conditions);
        Assert.Equal(new[] { "철손 42W, 동손 118W, 효율 93.1%" }, outline.Results);
        Assert.Equal(new[] { "A안 대비 손실 8% 감소" }, outline.Conclusions);
        Assert.Equal(new[] { @"D:\해석\모터A\run_0821.aedt" }, outline.Files);
        Assert.Equal(new[] { "열 해석 추가, 이번주 금요일까지" }, outline.Actions);
        Assert.True(outline.IsAnalysis);
        Assert.Equal(string.Empty, outline.Summary);
    }

    // ---- forgiveness ---------------------------------------------------------

    [Theory]
    [InlineData("결정: B안")]
    [InlineData("결정 : B안")]
    [InlineData("결정： B안")]
    [InlineData("  결정:   B안   ")]
    [InlineData("결정:B안")]
    public void TheColonIsForgiving(string line)
    {
        var outline = NoteOutline.Parse(line);
        Assert.Equal(new[] { "B안" }, outline.Decisions);
    }

    [Fact]
    public void LineOrderDoesNotMatter()
    {
        var outline = NoteOutline.Parse("다음: 후속\n결정: 하나\n요약문\n결정: 둘");

        Assert.Equal(new[] { "하나", "둘" }, outline.Decisions);
        Assert.Equal("요약문", outline.Summary);
    }

    [Fact]
    public void TheSameOpenerMayAppearMoreThanOnce()
    {
        var outline = NoteOutline.Parse("할일: 첫째\n할일: 둘째\n할일: 셋째");
        Assert.Equal(3, outline.Actions.Count);
    }

    [Fact]
    public void BlankLinesAreNotLines()
    {
        var outline = NoteOutline.Parse("\n\n요약\n\n\n결정: 하나\n\n");

        Assert.Equal(2, outline.Lines.Count);
    }

    [Fact]
    public void WindowsLineEndingsAreFine()
    {
        var outline = NoteOutline.Parse("요약\r\n결정: 하나\r\n");

        Assert.Equal("요약", outline.Summary);
        Assert.Equal(new[] { "하나" }, outline.Decisions);
    }

    // ---- what stays prose ----------------------------------------------------

    [Theory]
    [InlineData("참고: 지난주 자료")]          // not on the list
    [InlineData("결정 사항: 없음")]            // two words before the colon
    [InlineData("결정했다: B안")]              // not the bare word
    [InlineData("회의는 3시: 늦지 말 것")]     // a colon further along
    [InlineData("결정")]                       // no colon at all
    [InlineData("A결정: B안")]                 // not at the start
    public void AnythingElseIsText(string line)
    {
        var outline = NoteOutline.Parse(line);

        Assert.Single(outline.Lines);
        Assert.Equal(NoteLineKind.Text, outline.Lines[0].Kind);
        Assert.Equal(line.Trim(), outline.Lines[0].Text);
        Assert.False(outline.HasStructure);
    }

    [Fact]
    public void AnOpenerWithNothingAfterItIsLeftAsWritten()
    {
        // The person may mean to come back to it. Kept in sight rather than swallowed.
        var outline = NoteOutline.Parse("결정:\n결정:   ");

        Assert.Empty(outline.Decisions);
        Assert.Equal(new[] { "결정:", "결정:" }, outline.Text);
    }

    [Fact]
    public void AKeywordInsideTheTextIsNotAnOpener()
    {
        var outline = NoteOutline.Parse("오늘 결정: 없음, 다음 회의에서");

        Assert.Empty(outline.Decisions);
        Assert.Single(outline.Text);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   \n  \n")]
    public void NothingHasNoOutline(string? notes)
    {
        var outline = NoteOutline.Parse(notes);

        Assert.Same(NoteOutline.Empty, outline);
        Assert.Empty(outline.Lines);
        Assert.Equal(string.Empty, outline.Summary);
        Assert.False(outline.HasStructure);
    }

    // ---- the file line is the one the hand-over already writes ------------------

    [Fact]
    public void TheFileLineIsTheSameOneTheInboxWrites()
    {
        var notes = "파일: C:\\Users\\andrew\\Documents\\관리함\\보고서.pdf\n출처: FileBox 특별규칙 \"리포트 읽기\"";

        Assert.Equal(new[] { @"C:\Users\andrew\Documents\관리함\보고서.pdf" }, NoteOutline.Parse(notes).Files);
        Assert.Equal(new[] { "출처: FileBox 특별규칙 \"리포트 읽기\"" }, NoteOutline.Parse(notes).Text);
    }

    [Fact]
    public void ADrivLetterColonDoesNotConfuseTheFileLine()
    {
        Assert.Equal(new[] { @"C:\x.pdf" }, NoteOutline.Parse(@"파일:C:\x.pdf").Files);
    }
}

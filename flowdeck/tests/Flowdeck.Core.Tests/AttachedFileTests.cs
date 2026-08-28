using Flowdeck.Core.Services;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// Reading back the file a handed-over entry is about. The line is the sending side's half of
/// the bargain, so what counts as one is pinned here.
/// </summary>
public class AttachedFileTests
{
    [Fact]
    public void TheLineTheSendingSideWritesIsRead()
    {
        var notes = "파일: C:\\Users\\andrew\\Documents\\관리함\\2026_상반기_시장분석.pdf\n"
                  + "출처: FileBox 특별규칙 \"리포트 읽기\"";

        Assert.Equal(
            @"C:\Users\andrew\Documents\관리함\2026_상반기_시장분석.pdf",
            AttachedFile.PathIn(notes));
    }

    [Fact]
    public void ALineFurtherDownStillCounts()
    {
        // The sender is free to put a heading above it; only the line matters.
        Assert.Equal(
            @"C:\보고서.pdf",
            AttachedFile.PathIn("FileBox에서 보냄\n파일: C:\\보고서.pdf"));
    }

    [Fact]
    public void TheFirstOneWins()
    {
        Assert.Equal(
            @"C:\하나.pdf",
            AttachedFile.PathIn("파일: C:\\하나.pdf\n파일: C:\\둘.pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("장 볼 것 정리하기")]
    [InlineData("파일이 어디 갔더라")]      // starts with the word, but is not the line
    [InlineData("파일:")]                   // the line with nothing after it
    [InlineData("파일:    ")]
    public void AnythingElseNamesNoFile(string? notes) => Assert.Null(AttachedFile.PathIn(notes));

    [Fact]
    public void SurroundingSpaceIsNotPartOfThePath()
    {
        Assert.Equal(@"C:\보고서.pdf", AttachedFile.PathIn("  파일:   C:\\보고서.pdf   "));
    }

    [Fact]
    public void AWindowsLineEndingIsNotPartOfThePath()
    {
        Assert.Equal(@"C:\보고서.pdf", AttachedFile.PathIn("파일: C:\\보고서.pdf\r\n출처: FileBox"));
    }

    [Theory]
    [InlineData(@"C:\보고서.pdf")]
    [InlineData(@"C:\표.xlsx")]
    [InlineData(@"C:\사진.PNG")]
    [InlineData(@"C:\메모.txt")]
    [InlineData(@"C:\압축.zip")]
    public void AFileToReadIsOpenable(string path) => Assert.False(AttachedFile.IsRunnable(path));

    [Theory]
    [InlineData(@"C:\setup.exe")]
    [InlineData(@"C:\run.BAT")]
    [InlineData(@"C:\script.ps1")]
    [InlineData(@"C:\thing.vbs")]
    [InlineData(@"C:\바로가기.lnk")]
    [InlineData(@"C:\patch.reg")]
    [InlineData(@"C:\installer.msi")]
    public void AProgramIsNotOpenable(string path) => Assert.True(AttachedFile.IsRunnable(path));

    [Fact]
    public void NothingIsNotAProgram()
    {
        Assert.False(AttachedFile.IsRunnable(null));
        Assert.False(AttachedFile.IsRunnable(string.Empty));
        Assert.False(AttachedFile.IsRunnable(@"C:\확장자없음"));
    }
}

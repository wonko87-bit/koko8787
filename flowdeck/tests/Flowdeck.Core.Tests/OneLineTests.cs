using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// Text that arrives from somewhere other than a keyboard: a selection in a mail, a block
/// copied off a web page. It is shaped by the column it was set in, and none of that shape
/// means anything to the parser.
/// </summary>
public class OneLineTests
{
    private static readonly DateTime Now = new(2026, 8, 20, 9, 0, 0);

    private static NaturalLanguageParser Parser() => new();

    // ---- the fold itself ---------------------------------------------------

    [Fact]
    public void LineBreaksBecomeOneSpace()
    {
        Assert.Equal("다음주 화요일 김대리 미팅", NaturalLanguageParser.OneLine("다음주 화요일\n김대리 미팅"));
        Assert.Equal("한 줄", NaturalLanguageParser.OneLine("한\r\n줄"));
        Assert.Equal("한 줄", NaturalLanguageParser.OneLine("한\r줄"));
    }

    [Fact]
    public void ARunOfSpacingLeavesOneSpace()
    {
        Assert.Equal("보고서 제출", NaturalLanguageParser.OneLine("보고서 \t\n\n   제출"));
    }

    [Fact]
    public void NothingIsLeftHangingOffEitherEnd()
    {
        Assert.Equal("보고서", NaturalLanguageParser.OneLine("\n\n  보고서  \n\n"));
    }

    [Fact]
    public void TheSpacesThatCopiedTextIsFullOfCountAsSpaces()
    {
        // A non-breaking space looks like a space and is not one, and text lifted out of a
        // mail or a web page is full of them.
        Assert.Equal("오후 3시 회의", NaturalLanguageParser.OneLine("오후 3시 회의"));
    }

    [Fact]
    public void TheInvisibleOnesGoAltogether()
    {
        // Zero-width characters would sit inside a word and stop it matching anything, while
        // looking perfectly ordinary on screen.
        Assert.Equal("회의", NaturalLanguageParser.OneLine("﻿회​의"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   \n\t  ")]
    public void NothingFoldsToNothing(string? text) =>
        Assert.Equal(string.Empty, NaturalLanguageParser.OneLine(text));

    [Fact]
    public void TextThatIsAlreadyOneLineIsLeftAlone()
    {
        const string line = "!TD 내일 오후 3시 보고서 제출 #업무";
        Assert.Equal(line, NaturalLanguageParser.OneLine(line));
    }

    // ---- through the parser ------------------------------------------------

    [Fact]
    public void ADateSplitAcrossTwoLinesIsStillRead()
    {
        var entry = Parser().Parse("다음주 화요일\n오후 3시\n김대리 미팅", Now);

        Assert.Equal(new DateTime(2026, 8, 25, 15, 0, 0), entry.Start);
        Assert.Equal("김대리 미팅", entry.Title);
    }

    [Fact]
    public void MarkersAndTagsSurviveTheirOwnLines()
    {
        var entry = Parser().Parse("!TD\n보고서 제출\n#업무", Now);

        Assert.Equal(EntryTarget.Todo, entry.Target);
        Assert.True(entry.TargetWasExplicit);
        Assert.Equal("보고서 제출", entry.Title);
        Assert.Equal(new[] { "업무" }, entry.Tags);
    }

    [Fact]
    public void APastedBlockDoesNotLeaveItsSpacingInTheTitle()
    {
        var entry = Parser().Parse("  회의   준비\t\t자료  ", Now);

        Assert.Equal("회의 준비 자료", entry.Title);
    }

    [Fact]
    public void ANoteKeepsTheLineBreaksItWasGiven()
    {
        // The fold happens after the note is taken out. A note is prose, and prose is the one
        // place in the line where a line break was meant.
        var entry = Parser().Parse("보고서 제출 /* 첫째 줄\n둘째 줄 */", Now);

        Assert.Equal("보고서 제출", entry.Title);
        Assert.Equal("첫째 줄\n둘째 줄", entry.Notes);
    }

    [Fact]
    public void ANoteRunningToTheEndKeepsThemToo()
    {
        var entry = Parser().Parse("보고서 제출 /* 첫째 줄\n둘째 줄", Now);

        Assert.Equal("보고서 제출", entry.Title);
        Assert.Equal("첫째 줄\n둘째 줄", entry.Notes);
    }

    [Fact]
    public void WhatWasTypedIsStillRecordedAsItWasTyped()
    {
        const string pasted = "다음주 화요일\n김대리 미팅";
        Assert.Equal(pasted, Parser().Parse(pasted, Now).RawInput);
    }
}

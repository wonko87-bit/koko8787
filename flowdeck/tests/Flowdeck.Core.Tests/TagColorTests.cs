using Flowdeck.Core.Settings;
using Xunit;

namespace Flowdeck.Core.Tests;

/// <summary>
/// Each tag wears one colour, for good. Handed out in the order tags are first seen, kept
/// in settings, and the same whichever way the tag was typed.
/// </summary>
public class TagColorTests
{
    [Fact]
    public void TagsAreColouredInTheOrderTheyAreFirstSeen()
    {
        var settings = new AppSettings();

        Assert.Equal(0, settings.TagColor("모터A", out var first));
        Assert.Equal(1, settings.TagColor("보고", out var second));
        Assert.Equal(2, settings.TagColor("리뷰", out var third));

        Assert.True(first);
        Assert.True(second);
        Assert.True(third);
    }

    [Fact]
    public void ATagKeepsItsColour()
    {
        var settings = new AppSettings();
        settings.TagColor("모터A", out _);
        settings.TagColor("보고", out _);

        Assert.Equal(0, settings.TagColor("모터A", out var assigned));
        Assert.False(assigned);
    }

    [Fact]
    public void CaseDoesNotMakeANewTag()
    {
        var settings = new AppSettings();
        settings.TagColor("Motor", out _);

        Assert.Equal(0, settings.TagColor("motor", out var assigned));
        Assert.False(assigned);
        Assert.Single(settings.TagColors);
    }

    [Fact]
    public void AfterTenTheColoursComeRoundAgain()
    {
        var settings = new AppSettings();
        for (var i = 0; i < AppSettings.TagColorCount; i++) settings.TagColor("태그" + i, out _);

        Assert.Equal(0, settings.TagColor("열한번째", out _));
    }

    [Fact]
    public void AColourSetByHandIsRespected()
    {
        var settings = new AppSettings { TagColors = { ["급함"] = 3 } };

        Assert.Equal(3, settings.TagColor("급함", out var assigned));
        Assert.False(assigned);

        // The next new tag counts from how many are known, not from the highest number.
        Assert.Equal(1, settings.TagColor("새것", out _));
    }

    [Fact]
    public void AnOutOfRangeNumberInTheFileIsFoldedIntoThePalette()
    {
        var settings = new AppSettings { TagColors = { ["큰수"] = 23, ["음수"] = -1 } };

        Assert.Equal(3, settings.TagColor("큰수", out _));
        Assert.Equal(9, settings.TagColor("음수", out _));
    }

    [Fact]
    public void TheColoursSurviveTheSettingsFile()
    {
        var settings = new AppSettings();
        settings.TagColor("모터A", out _);
        settings.TagColor("보고", out _);

        var path = Path.Combine(Path.GetTempPath(), "flowdeck-tagcolors-" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            settings.SaveTo(path);
            var loaded = AppSettings.LoadFrom(path);

            Assert.Equal(1, loaded.TagColor("보고", out var assigned));
            Assert.False(assigned);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void NothingHasNoColour(string? tag)
    {
        var settings = new AppSettings();

        Assert.Equal(-1, settings.TagColor(tag!, out var assigned));
        Assert.False(assigned);
        Assert.Empty(settings.TagColors);
    }
}

using System.Text.Json;
using Flowdeck.Core.Settings;
using Xunit;

namespace Flowdeck.Core.Tests;

public class SettingsTests
{
    /// <summary>
    /// A brand new settings object has no stored window positions. It still has to
    /// serialise: this used to hold double.NaN, which System.Text.Json refuses to write,
    /// so the first save of a fresh profile threw and took the app down with it.
    /// </summary>
    [Fact]
    public void DefaultSettingsSerialise()
    {
        var settings = new AppSettings();

        var json = JsonSerializer.Serialize(settings);

        Assert.Contains("WidgetLeft", json);
        Assert.DoesNotContain("NaN", json);
    }

    [Fact]
    public void DefaultSettingsSurviveASaveAndLoad()
    {
        var directory = Path.Combine(Path.GetTempPath(), "flowdeck-settings-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "settings.json");

        try
        {
            new AppSettings().SaveTo(path);
            var reloaded = AppSettings.LoadFrom(path);

            Assert.Null(reloaded.WidgetLeft);
            Assert.Null(reloaded.WidgetTop);
            Assert.False(reloaded.AgendaEventsPlacement.HasPosition);
            Assert.Equal(DayOfWeek.Monday, reloaded.FirstDayOfWeek);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PositionsRoundTripOnceTheyAreSet()
    {
        var directory = Path.Combine(Path.GetTempPath(), "flowdeck-settings-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "settings.json");

        try
        {
            var settings = new AppSettings { WidgetLeft = 120.5, WidgetTop = 40 };
            settings.AgendaTodosPlacement.Left = -8;
            settings.AgendaTodosPlacement.Top = 15;
            settings.SaveTo(path);

            var reloaded = AppSettings.LoadFrom(path);

            Assert.Equal(120.5, reloaded.WidgetLeft);
            Assert.Equal(40, reloaded.WidgetTop);
            Assert.True(reloaded.AgendaTodosPlacement.HasPosition);
            Assert.Equal(-8, reloaded.AgendaTodosPlacement.Left);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void AnUnreadableSettingsFileFallsBackToDefaults()
    {
        var directory = Path.Combine(Path.GetTempPath(), "flowdeck-settings-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "settings.json");

        try
        {
            File.WriteAllText(path, "{ not json");

            Assert.Equal("Ctrl+Alt+Space", AppSettings.LoadFrom(path).QuickAddHotkey);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}

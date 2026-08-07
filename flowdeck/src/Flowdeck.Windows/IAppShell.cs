using Flowdeck.Core.Services;
using Flowdeck.Core.Settings;

namespace Flowdeck.Windows;

/// <summary>
/// What the windows are allowed to ask of the application as a whole.
/// Keeping it to an interface stops the views from reaching into <see cref="App"/> directly.
/// </summary>
public interface IAppShell
{
    AppSettings Settings { get; }

    WorkspaceRepository Repository { get; }

    /// <summary>Folder holding workspace.json and settings.json.</summary>
    string DataFolder { get; }

    void ApplyTheme(AppTheme theme);

    /// <summary>Pushes opacity and pin-mode changes into the live widget.</summary>
    void ApplyWidgetSettings();

    /// <summary>Applies parser-affecting settings such as the keyword and afternoon rules.</summary>
    void ApplyParserSettings();

    /// <summary>
    /// Re-claims the configured hot keys. Returns a message describing any that could
    /// not be taken, or an empty string when all of them registered.
    /// </summary>
    string ReapplyHotKeys();

    void SaveSettings();

    void ShowWidget();

    void HideWidget();

    void ShowQuickAdd();

    /// <summary>Opens the accumulated calendar list, or closes it if it is already up.</summary>
    void ToggleEventsAgenda();

    /// <summary>Opens the accumulated todo list, or closes it if it is already up.</summary>
    void ToggleTodosAgenda();
}

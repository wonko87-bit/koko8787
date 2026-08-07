using System.Drawing;
using System.Windows;
using WinForms = System.Windows.Forms;

namespace Flowdeck.Windows.Services;

/// <summary>
/// The notification-area icon and its menu. WPF has no tray icon of its own, so this
/// is the one place the app borrows from Windows Forms.
/// </summary>
public sealed class TrayIconController : IDisposable
{
    private readonly WinForms.NotifyIcon _icon;
    private readonly WinForms.ToolStripMenuItem _widgetItem;

    public TrayIconController()
    {
        var menu = new WinForms.ContextMenuStrip { ShowImageMargin = false };

        var quickAdd = new WinForms.ToolStripMenuItem("빠른 입력");
        quickAdd.Click += (_, _) => QuickAddRequested?.Invoke(this, EventArgs.Empty);
        quickAdd.Font = new Font(quickAdd.Font, FontStyle.Bold);

        _widgetItem = new WinForms.ToolStripMenuItem("위젯 표시");
        _widgetItem.Click += (_, _) => WidgetToggleRequested?.Invoke(this, EventArgs.Empty);

        var settings = new WinForms.ToolStripMenuItem("설정");
        settings.Click += (_, _) => SettingsRequested?.Invoke(this, EventArgs.Empty);

        var exit = new WinForms.ToolStripMenuItem("종료");
        exit.Click += (_, _) => ExitRequested?.Invoke(this, EventArgs.Empty);

        menu.Items.Add(quickAdd);
        menu.Items.Add(_widgetItem);
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add(settings);
        menu.Items.Add(exit);

        _icon = new WinForms.NotifyIcon
        {
            Icon = LoadIcon(),
            Text = "Flowdeck",
            Visible = true,
            ContextMenuStrip = menu,
        };

        _icon.DoubleClick += (_, _) => WidgetToggleRequested?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? QuickAddRequested;

    public event EventHandler? WidgetToggleRequested;

    public event EventHandler? SettingsRequested;

    public event EventHandler? ExitRequested;

    /// <summary>Keeps the menu label matching the widget's actual state.</summary>
    public void SetWidgetVisible(bool visible) => _widgetItem.Text = visible ? "위젯 숨기기" : "위젯 표시";

    public void Notify(string title, string message) =>
        _icon.ShowBalloonTip(6000, title, message, WinForms.ToolTipIcon.None);

    private static Icon LoadIcon()
    {
        var uri = new Uri("pack://application:,,,/Assets/flowdeck.ico", UriKind.Absolute);
        var resource = Application.GetResourceStream(uri);

        // SystemIcons.Application is only a fallback for a build that lost the asset.
        if (resource is null) return SystemIcons.Application;

        using var stream = resource.Stream;
        return new Icon(stream, new Size(16, 16));
    }

    public void Dispose()
    {
        _icon.Visible = false;
        _icon.Dispose();
    }
}

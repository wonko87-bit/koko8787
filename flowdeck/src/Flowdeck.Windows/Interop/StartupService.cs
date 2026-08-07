using System.Diagnostics;
using Microsoft.Win32;

namespace Flowdeck.Windows.Interop;

/// <summary>
/// Registers the app under the per-user Run key so it comes back after a reboot.
/// Per-user rather than machine-wide, so it never needs elevation.
/// </summary>
public static class StartupService
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "Flowdeck";

    public static bool IsEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: false);
            return key?.GetValue(ValueName) is string value && value.Length > 0;
        }
        catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
        {
            return false;
        }
    }

    /// <summary>Returns false when the registry refused the write, so the UI can stay honest.</summary>
    public static bool Set(bool enabled)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKey, writable: true);
            if (key is null) return false;

            if (enabled)
            {
                var path = ExecutablePath();
                if (path is null) return false;
                key.SetValue(ValueName, $"\"{path}\"");
            }
            else
            {
                key.DeleteValue(ValueName, throwOnMissingValue: false);
            }

            return true;
        }
        catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
        {
            return false;
        }
    }

    private static string? ExecutablePath()
    {
        // Environment.ProcessPath is the published .exe; the module name is the fallback
        // for the rare host that leaves it null.
        var path = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(path)) return path;

        return Process.GetCurrentProcess().MainModule?.FileName;
    }
}

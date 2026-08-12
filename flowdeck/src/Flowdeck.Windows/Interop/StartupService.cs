using System.Diagnostics;
using Microsoft.Win32;

namespace Flowdeck.Windows.Interop;

/// <summary>Why the app is not actually starting with Windows, when the box says it should.</summary>
public enum StartupState
{
    /// <summary>Not registered at all.</summary>
    Off,

    /// <summary>Registered, approved, and pointing at this copy of the app.</summary>
    On,

    /// <summary>
    /// Registered, but pointing at an executable somewhere else — usually the previous
    /// build, moved or replaced since. Windows runs that path and finds nothing.
    /// </summary>
    StalePath,

    /// <summary>
    /// Registered, but Windows has been told to ignore it — Task Manager's Startup tab, or
    /// Settings → Apps → Startup. The Run entry is still there, doing nothing.
    /// </summary>
    BlockedByWindows,
}

/// <summary>
/// Registers the app under the per-user Run key so it comes back after a reboot.
/// Per-user rather than machine-wide, so it never needs elevation.
///
/// The Run key alone does not settle the question. Windows keeps a veto of its own in
/// StartupApproved, written whenever the entry is switched off in Task Manager or Settings,
/// and it wins silently. And the entry holds a full path, so a build replaced in a different
/// folder leaves the old one pointing at nothing. Both look identical from the Run key —
/// registered — which is why reading only that would have the checkbox lie.
/// </summary>
public static class StartupService
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

    /// <summary>Explorer's own record of which Run entries the user has switched off.</summary>
    private const string ApprovedKey =
        @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";

    private const string ValueName = "Flowdeck";

    public static bool IsEnabled() => Current() == StartupState.On;

    /// <summary>What is actually going to happen at the next sign-in.</summary>
    public static StartupState Current()
    {
        var registered = RegisteredPath();
        if (registered is null) return StartupState.Off;

        if (IsBlockedByWindows()) return StartupState.BlockedByWindows;

        var current = ExecutablePath();
        if (current is not null && !string.Equals(registered, current, StringComparison.OrdinalIgnoreCase))
        {
            return StartupState.StalePath;
        }

        return StartupState.On;
    }

    /// <summary>The executable the Run entry points at, unquoted. Null when there is no entry.</summary>
    public static string? RegisteredPath()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: false);
            if (key?.GetValue(ValueName) is not string value) return null;

            var path = value.Trim().Trim('"');
            return path.Length == 0 ? null : path;
        }
        catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
        {
            return null;
        }
    }

    /// <summary>
    /// Rewrites the path when it has drifted, and only then.
    ///
    /// Called at every start so that replacing the build — a new folder, a new download —
    /// does not quietly unregister the app. Deliberately leaves a Windows veto alone: that
    /// one was somebody's decision, and undoing it behind their back every launch would be
    /// the app arguing with the user once a day.
    /// </summary>
    public static bool RepairIfStale()
    {
        if (Current() != StartupState.StalePath) return false;

        return Write(ExecutablePath());
    }

    /// <summary>
    /// Registers or unregisters. Returns false when the registry refused the write, so the
    /// UI can stay honest.
    ///
    /// Switching it on clears any Windows veto, because ticking the box here is the user
    /// saying so plainly — which is the one moment overriding that decision is right.
    /// </summary>
    public static bool Set(bool enabled)
    {
        if (!enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RunKey, writable: true);
                key?.DeleteValue(ValueName, throwOnMissingValue: false);
                return key is not null;
            }
            catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
            {
                return false;
            }
        }

        if (!Write(ExecutablePath())) return false;

        ClearWindowsBlock();
        return true;
    }

    private static bool Write(string? path)
    {
        if (path is null) return false;

        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKey, writable: true);
            if (key is null) return false;

            // Quoted: the path holds spaces on any ordinary install.
            key.SetValue(ValueName, $"\"{path}\"");
            return true;
        }
        catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
        {
            return false;
        }
    }

    /// <summary>
    /// True when Explorer has been told to ignore the Run entry. The record is twelve bytes
    /// whose first one carries the flag; the low bit set means disabled.
    /// </summary>
    private static bool IsBlockedByWindows()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(ApprovedKey, writable: false);
            if (key?.GetValue(ValueName) is not byte[] { Length: > 0 } record) return false;

            return (record[0] & 1) == 1;
        }
        catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
        {
            return false;
        }
    }

    /// <summary>
    /// Removes the veto. Explorer treats a missing record as approved and writes a fresh one
    /// itself, which is safer than guessing the twelve bytes it wants.
    /// </summary>
    private static void ClearWindowsBlock()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(ApprovedKey, writable: true);
            key?.DeleteValue(ValueName, throwOnMissingValue: false);
        }
        catch (Exception e) when (e is UnauthorizedAccessException or System.Security.SecurityException)
        {
            // Nothing to do about it, and it is not worth failing the whole change over.
        }
    }

    private static string? ExecutablePath()
    {
        // Environment.ProcessPath is the published .exe; the module name is the fallback
        // for the rare host that leaves it null.
        var path = Environment.ProcessPath;
        if (string.IsNullOrEmpty(path)) path = Process.GetCurrentProcess().MainModule?.FileName;

        return string.IsNullOrEmpty(path) ? null : path;
    }
}

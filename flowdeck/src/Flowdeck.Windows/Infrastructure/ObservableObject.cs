using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Flowdeck.Windows.Infrastructure;

/// <summary>
/// Minimal change notification. Hand-rolled on purpose: the app has no NuGet
/// dependencies, so it restores and builds on a machine with no package feed.
/// </summary>
public abstract class ObservableObject : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected void Raise([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;

        field = value;
        Raise(propertyName);
        return true;
    }
}

using System.Collections.ObjectModel;
using System.Windows.Input;
using Flowdeck.Core.Parsing;
using Flowdeck.Windows.Infrastructure;

namespace Flowdeck.Windows.ViewModels;

/// <summary>One person in the address book, as shown in the editor.</summary>
public sealed class ContactRow : ObservableObject
{
    private readonly Action _changed;
    private string _name;
    private string _address;

    public ContactRow(string name, string address, Action changed, Action<ContactRow> remove)
    {
        _name = name;
        _address = address;
        _changed = changed;
        RemoveCommand = new RelayCommand(() => remove(this));
    }

    /// <summary>What is typed after the <c>@</c>.</summary>
    public string Name
    {
        get => _name;
        set
        {
            if (Set(ref _name, value)) Revalidate();
        }
    }

    public string Address
    {
        get => _address;
        set
        {
            if (Set(ref _address, value)) Revalidate();
        }
    }

    public ICommand RemoveCommand { get; }

    /// <summary>False when this row would never be found, so the editor can mark it.</summary>
    public bool IsValid => ContactBook.IsUsableName(_name) && ContactBook.IsAddress(_address);

    public bool IsNotValid => !IsValid;

    private void Revalidate()
    {
        Raise(nameof(IsValid));
        Raise(nameof(IsNotValid));
        _changed();
    }
}

/// <summary>
/// Backs the address-book window: the same list that lives in settings.json, editable
/// without opening settings.json. Every change writes straight through, matching how the
/// rest of the settings behave — there is no save button to forget.
/// </summary>
public sealed class ContactsViewModel : ObservableObject
{
    private readonly ContactBook _book;
    private readonly Action _persist;
    private readonly RelayCommand _add;

    private string _newName = string.Empty;
    private string _newAddress = string.Empty;
    private string _status = string.Empty;

    public ContactsViewModel(ContactBook book, Action persist)
    {
        _book = book;
        _persist = persist;

        _add = new RelayCommand(Add, () => NewName.Length > 0 || NewAddress.Length > 0);

        foreach (var pair in book.Aliases.OrderBy(p => p.Key, StringComparer.CurrentCulture))
        {
            Contacts.Add(NewRow(pair.Key, pair.Value));
        }

        RaiseCounts();
    }

    public ObservableCollection<ContactRow> Contacts { get; } = new();

    public ICommand AddCommand => _add;

    public string NewName
    {
        get => _newName;
        set
        {
            if (Set(ref _newName, value)) _add.RaiseCanExecuteChanged();
        }
    }

    public string NewAddress
    {
        get => _newAddress;
        set
        {
            if (Set(ref _newAddress, value)) _add.RaiseCanExecuteChanged();
        }
    }

    /// <summary>The last thing that happened, or why the last thing did not happen.</summary>
    public string Status
    {
        get => _status;
        set => Set(ref _status, value);
    }

    public bool IsEmpty => Contacts.Count == 0;

    public string CountLabel => Contacts.Count == 0 ? string.Empty : $"{Contacts.Count}명";

    /// <summary>
    /// Adds the pending row after checking it would actually be found. A name with a space
    /// in it is the trap worth catching here: it saves without complaint and then never
    /// matches anything, because "@김 대리" stops at the space.
    /// </summary>
    public void Add()
    {
        var name = NewName.Trim().TrimStart('@');
        var address = NewAddress.Trim();

        if (!ContactBook.IsUsableName(name))
        {
            Status = name.Length == 0
                ? "이름을 입력하세요"
                : "이름에는 공백이나 특수문자를 쓸 수 없습니다";
            return;
        }

        if (!ContactBook.IsAddress(address))
        {
            Status = address.Length == 0 ? "메일 주소를 입력하세요" : "메일 주소 형식이 아닙니다";
            return;
        }

        var existing = Find(name);
        if (existing is not null)
        {
            // Replacing beats a second row with the same name, which would resolve to
            // whichever one happened to be written last.
            existing.Address = address;
            Status = $"{name} 의 주소를 바꿨습니다";
        }
        else
        {
            Contacts.Add(NewRow(name, address));
            Status = $"{name} 을(를) 추가했습니다";
        }

        NewName = string.Empty;
        NewAddress = string.Empty;
        Commit();
    }

    public void Remove(ContactRow row)
    {
        if (!Contacts.Remove(row)) return;

        Status = row.Name.Length > 0 ? $"{row.Name} 을(를) 지웠습니다" : "항목을 지웠습니다";
        Commit();
    }

    private ContactRow? Find(string name) =>
        Contacts.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));

    private ContactRow NewRow(string name, string address) => new(name, address, Commit, Remove);

    /// <summary>
    /// Rewrites the book from the rows and saves. Incomplete rows are simply left out
    /// rather than rejected, so a half-typed line never blocks the rest from being kept.
    /// </summary>
    private void Commit()
    {
        _book.Aliases.Clear();

        foreach (var row in Contacts.Where(c => c.IsValid))
        {
            _book.Aliases[row.Name.Trim()] = row.Address.Trim();
        }

        RaiseCounts();
        _persist();
    }

    private void RaiseCounts()
    {
        Raise(nameof(IsEmpty));
        Raise(nameof(CountLabel));
    }
}

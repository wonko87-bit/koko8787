using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Flowdeck.Core.Models;

namespace Flowdeck.Core.Export;

/// <summary>
/// A batch of entries on its way from one installation to another.
///
/// Flowdeck's own shape rather than iCalendar, because .ics carries an event well and a
/// todo badly: tags, priority, the repeat description and the original typed line all
/// survive here and would not survive there.
/// </summary>
public sealed class TransferArchive
{
    public const string FormatName = "flowdeck.transfer";

    public const int CurrentVersion = 1;

    /// <summary>Names the shape, so a JSON file that is something else is caught on sight.</summary>
    public string Format { get; set; } = FormatName;

    public int Version { get; set; } = CurrentVersion;

    public DateTimeOffset ExportedAt { get; set; }

    public List<TodoItem> Todos { get; set; } = new();

    public List<CalendarEvent> Events { get; set; } = new();

    [JsonIgnore]
    public int Count => Todos.Count + Events.Count;
}

/// <summary>
/// Reads and writes <see cref="TransferArchive"/>. Pure text in, pure text out — no file
/// system — so the transfer format can be tested and reused wherever the app is running.
/// </summary>
public static class TransferFile
{
    /// <summary>
    /// The extension to save under.
    ///
    /// Plain .txt because that is what gets through. A corporate messenger will carry a
    /// text file and refuse a .json or a .csv without explaining itself, and a file that
    /// cannot be sent is no use however well-formed it is. The contents are still JSON —
    /// only the name had to give way.
    /// </summary>
    public const string Extension = ".txt";

    /// <summary>
    /// Separates the part written for a person from the part written for the app. Anyone
    /// previewing the file in a chat window sees what is in it; the reader skips past.
    /// </summary>
    private const string Marker = "--- 여기서부터는 앱이 읽는 부분입니다. 지우지 마세요 ---";

    private static readonly CultureInfo Korean = CultureInfo.GetCultureInfo("ko-KR");

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static string Write(
        IEnumerable<TodoItem> todos,
        IEnumerable<CalendarEvent> events,
        DateTimeOffset exportedAt)
    {
        var archive = new TransferArchive
        {
            ExportedAt = exportedAt,

            // Copies, because stripping the external link below must not reach the records
            // still being used by the app that is doing the exporting.
            Todos = todos.Select(Strip).ToList(),
            Events = events.Select(Strip).ToList(),
        };

        var text = new StringBuilder();
        text.Append("Flowdeck 내보내기 · ")
            .Append(exportedAt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture))
            .Append(" · ")
            .Append(archive.Count)
            .AppendLine("건")
            .AppendLine();

        foreach (var todo in archive.Todos) text.Append("- [할일] ").AppendLine(Summarise(todo.Title, todo.DueAt, todo.HasTime, todo.Tags));
        foreach (var e in archive.Events) text.Append("- [일정] ").AppendLine(Summarise(e.Title, e.Start, !e.IsAllDay, e.Tags));

        text.AppendLine().AppendLine(Marker);
        text.Append(JsonSerializer.Serialize(archive, Options));

        return text.ToString();
    }

    /// <summary>One entry as a person would read it, for the human half of the file.</summary>
    private static string Summarise(string title, DateTime? at, bool hasTime, List<string> tags)
    {
        var line = new StringBuilder(title);

        if (at.HasValue)
        {
            line.Append(" · ").Append(at.Value.ToString(hasTime ? "M월 d일 HH:mm" : "M월 d일", Korean));
        }

        if (tags.Count > 0) line.Append(' ').Append(string.Join(" ", tags.Select(t => "#" + t)));

        return line.ToString();
    }

    /// <summary>
    /// Parses a transfer file, or explains in one sentence why it is not one. Every failure
    /// mode ends up in front of a person who chose a file, so each says which file problem
    /// it was rather than surfacing a parser message.
    /// </summary>
    public static TransferArchive Read(string contents)
    {
        TransferArchive? archive;

        // Everything before the marker is for the reader's eyes, not the parser's. A file
        // written before the marker existed has none, and is JSON from its first character.
        var marker = contents?.IndexOf(Marker, StringComparison.Ordinal) ?? -1;
        var json = marker >= 0 ? contents!.Substring(marker + Marker.Length) : contents ?? string.Empty;

        try
        {
            archive = JsonSerializer.Deserialize<TransferArchive>(json, Options);
        }
        catch (JsonException)
        {
            throw new FormatException("파일을 읽을 수 없습니다. Flowdeck 내보내기 파일이 맞는지 확인해 주세요.");
        }

        if (archive is null || !string.Equals(archive.Format, TransferArchive.FormatName, StringComparison.Ordinal))
        {
            throw new FormatException("Flowdeck 내보내기 파일이 아닙니다.");
        }

        if (archive.Version > TransferArchive.CurrentVersion)
        {
            throw new FormatException($"더 새로운 버전({archive.Version})에서 만든 파일입니다. Flowdeck을 최신으로 올려 주세요.");
        }

        // An explicit "Todos": null in the file leaves the property null rather than empty,
        // and every caller past this point is entitled to iterate it.
        archive.Todos ??= new List<TodoItem>();
        archive.Events ??= new List<CalendarEvent>();

        return archive;
    }

    /// <summary>
    /// An external link names an entry inside one machine's Outlook, which the machine at
    /// the other end has never heard of. Carrying it across would be worse than useless:
    /// the receiving copy would look like it had already been sent, and would never be.
    /// </summary>
    private static TodoItem Strip(TodoItem todo)
    {
        var copy = todo.Clone();
        copy.ExternalLink = null;
        return copy;
    }

    private static CalendarEvent Strip(CalendarEvent calendarEvent)
    {
        var copy = calendarEvent.Clone();
        copy.ExternalLink = null;
        return copy;
    }
}

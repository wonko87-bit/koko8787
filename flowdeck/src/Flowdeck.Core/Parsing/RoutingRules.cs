using Flowdeck.Core.Models;

namespace Flowdeck.Core.Parsing;

/// <summary>Where a line goes when it says nothing about where it belongs.</summary>
public enum TargetDefault
{
    /// <summary>Work it out from the words, falling back to both. The original behaviour.</summary>
    Auto = 0,

    Todo = 1,

    Calendar = 2,

    Both = 3,
}

/// <summary>
/// Decides whether a line of input belongs on the calendar, on the todo list, or on both.
///
/// Four mechanisms, in strict order of precedence:
///
///   1. A marker at the head or the tail of the line. Symbols (<c>!TD</c>, <c>!CD</c>) and
///      plain words ("할일", "일정") work the same way and always win outright. A marker only
///      counts when it stands alone, so "일정관리 앱 알아보기" is left untouched.
///   2. A <c>#할일</c> / <c>#일정</c> hashtag, which may sit anywhere in the line.
///   3. <see cref="DefaultTarget"/>, when the user has named one.
///   4. Failing all three, a keyword sweep over the text. Words like "회의" pull towards the
///      calendar, words like "제출" pull towards the todo list. Switched off with
///      <see cref="UseKeywordHints"/>.
///
/// When nothing fires — or when both keyword lists match — the entry goes to both places,
/// which is the shipped default.
/// </summary>
public sealed class RoutingRules
{
    /// <summary>Markers that force calendar-only, matched case-insensitively.</summary>
    public List<string> CalendarMarkers { get; set; } = new()
    {
        "!CD", "!cal", "!일정", "!캘", "/c", "일정", "캘린더", "달력",
    };

    /// <summary>Markers that force todo-only, matched case-insensitively.</summary>
    public List<string> TodoMarkers { get; set; } = new()
    {
        "!TD", "!todo", "!할일", "!투두", "/t", "할일", "투두", "todo",
    };

    /// <summary>Markers that force both, for overriding a keyword rule.</summary>
    public List<string> BothMarkers { get; set; } = new()
    {
        "!BD", "!both", "!둘다", "!모두", "/b", "둘다",
    };

    /// <summary>
    /// Where an unmarked line goes. Someone who lives in one half of the app can set it
    /// there and stop typing the marker for the common case, leaving <c>!TD</c> / <c>!CD</c>
    /// — or Ctrl+1/2/3 in the capture window — for the exception.
    /// </summary>
    public TargetDefault DefaultTarget { get; set; } = TargetDefault.Auto;

    /// <summary>
    /// When false, only explicit markers route an entry and everything else goes to both places.
    /// Turn this off to get strictly predictable behaviour. Ignored once
    /// <see cref="DefaultTarget"/> names somewhere, which is a stronger statement of the same wish.
    /// </summary>
    public bool UseKeywordHints { get; set; } = true;

    /// <summary>Also copy this entry to Outlook.</summary>
    public List<string> ExternalMarkers { get; set; } = new()
    {
        "!OL", "!outlook", "!아웃룩",
    };

    /// <summary>Keep this entry local, overriding the always-push setting for one line.</summary>
    public List<string> LocalOnlyMarkers { get; set; } = new()
    {
        "!NOL", "!local", "!로컬",
    };

    /// <summary>Open the Teams new-meeting form for this entry, filled in but not sent.</summary>
    public List<string> TeamsMarkers { get; set; } = new()
    {
        "!TM", "!teams", "!팀즈",
    };

    /// <summary>
    /// Opens a note. Everything from here to <see cref="NoteClose"/> is kept verbatim and
    /// taken out of the line; with no closer, the note simply runs to the end.
    /// </summary>
    public string NoteOpen { get; set; } = "/*";

    /// <summary>Closes a note and hands the rest of the line back to the parser.</summary>
    public string NoteClose { get; set; } = "*/";

    /// <summary>Typed inside a note to break the line, since a text box has only one.</summary>
    public string NoteLineBreak { get; set; } = @"\n";

    /// <summary>Words that suggest a block of time in the day: something you attend.</summary>
    public List<string> CalendarKeywords { get; set; } = new()
    {
        "회의", "미팅", "약속", "세미나", "워크샵", "워크숍", "면담", "인터뷰", "면접",
        "발표", "행사", "출장", "강의", "수업", "진료", "예약", "컨퍼런스", "브리핑",
        "회식", "점심", "저녁식사", "생일", "기념일", "휴가", "연차", "반차", "party",
        "meeting", "appointment", "call",
    };

    /// <summary>Words that suggest a piece of work: something you complete.</summary>
    public List<string> TodoKeywords { get; set; } = new()
    {
        "제출", "작성", "정리", "확인", "검토", "구매", "구입", "결제", "신청", "등록",
        "처리", "마감", "준비", "예습", "복습", "청소", "빨래", "설거지", "전화하기",
        "보내기", "답장", "회신", "리뷰", "수정", "배포", "테스트", "버그", "todo",
        "사기", "읽기", "쓰기", "챙기기",
    };

    /// <summary>
    /// Finds an explicit marker and returns the text with it removed.
    /// Returns null when the input carries no marker at all.
    /// </summary>
    public EntryTarget? FindMarker(string input, out string remainder)
    {
        remainder = input.Trim();

        foreach (var (markers, target) in Ordered())
        {
            foreach (var marker in markers)
            {
                if (TryStripMarker(remainder, marker, out var stripped))
                {
                    remainder = stripped;
                    return target;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Finds an <c>!OL</c> / <c>!NOL</c> marker and returns the text with it removed.
    ///
    /// This sits at right angles to the todo/calendar decision rather than alongside it:
    /// Outlook is somewhere an entry is also copied to, not a third kind of entry. So the
    /// two combine — the routing markers pick which Outlook folder, this picks whether to
    /// send at all. Returns null when the line says nothing either way.
    /// </summary>
    public bool? FindExternalMarker(string input, out string remainder)
    {
        remainder = input.Trim();

        foreach (var marker in ExternalMarkers)
        {
            if (TryStripMarker(remainder, marker, out var pushed))
            {
                remainder = pushed;
                return true;
            }
        }

        foreach (var marker in LocalOnlyMarkers)
        {
            if (TryStripMarker(remainder, marker, out var local))
            {
                remainder = local;
                return false;
            }
        }

        return null;
    }

    /// <summary>
    /// Finds a <c>!TM</c> marker and returns the text with it removed.
    ///
    /// Another axis again: Teams is neither a kind of entry nor a place to copy one, but a
    /// window to open with the entry already typed into it. Nothing is sent — the user
    /// still presses Teams' own save button, which is the point of going this way round.
    /// </summary>
    public bool FindTeamsMarker(string input, out string remainder)
    {
        remainder = input.Trim();

        foreach (var marker in TeamsMarkers)
        {
            if (TryStripMarker(remainder, marker, out var stripped))
            {
                remainder = stripped;
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Decides where a line goes when it carried no marker of its own.
    ///
    /// A named <see cref="DefaultTarget"/> is taken at its word and the keyword sweep is
    /// skipped: having said "everything is a todo unless I say otherwise", the user should
    /// not then find "회의" quietly filed as an event. Only Auto reaches the keywords, where
    /// a tie — no keywords, or keywords from both lists — means both.
    /// </summary>
    public EntryTarget Classify(string text)
    {
        switch (DefaultTarget)
        {
            case TargetDefault.Todo: return EntryTarget.Todo;
            case TargetDefault.Calendar: return EntryTarget.Calendar;
            case TargetDefault.Both: return EntryTarget.Both;
        }

        if (!UseKeywordHints) return EntryTarget.Both;

        var calendarHit = CalendarKeywords.Any(k => Contains(text, k));
        var todoHit = TodoKeywords.Any(k => Contains(text, k));

        if (calendarHit && !todoHit) return EntryTarget.Calendar;
        if (todoHit && !calendarHit) return EntryTarget.Todo;
        return EntryTarget.Both;
    }

    private IEnumerable<(List<string> Markers, EntryTarget Target)> Ordered()
    {
        yield return (BothMarkers, EntryTarget.Both);
        yield return (CalendarMarkers, EntryTarget.Calendar);
        yield return (TodoMarkers, EntryTarget.Todo);
    }

    /// <summary>
    /// A marker counts only when it stands alone at the head or the tail of the line,
    /// so a word that merely starts with the same letters is left untouched.
    /// </summary>
    private static bool TryStripMarker(string text, string marker, out string remainder)
    {
        remainder = text;
        if (marker.Length == 0) return false;

        if (text.StartsWith(marker, StringComparison.OrdinalIgnoreCase))
        {
            var rest = text.Substring(marker.Length);
            if (rest.Length == 0 || char.IsWhiteSpace(rest[0]))
            {
                remainder = rest.Trim();
                return true;
            }
        }

        if (text.EndsWith(marker, StringComparison.OrdinalIgnoreCase))
        {
            var head = text.Substring(0, text.Length - marker.Length);
            if (head.Length > 0 && char.IsWhiteSpace(head[head.Length - 1]))
            {
                remainder = head.Trim();
                return true;
            }
        }

        return false;
    }

    private static bool Contains(string haystack, string needle) =>
        haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
}

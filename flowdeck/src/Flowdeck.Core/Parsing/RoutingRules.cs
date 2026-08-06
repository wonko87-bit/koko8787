using Flowdeck.Core.Models;

namespace Flowdeck.Core.Parsing;

/// <summary>
/// Decides whether a line of input belongs on the calendar, on the todo list, or on both.
///
/// Three mechanisms, in strict order of precedence:
///
///   1. A marker at the head or the tail of the line. Symbols (<c>!TD</c>, <c>!CD</c>) and
///      plain words ("할일", "일정") work the same way and always win outright. A marker only
///      counts when it stands alone, so "일정관리 앱 알아보기" is left untouched.
///   2. A <c>#할일</c> / <c>#일정</c> hashtag, which may sit anywhere in the line.
///   3. Failing both, a keyword sweep over the text. Words like "회의" pull towards the
///      calendar, words like "제출" pull towards the todo list. Switched off with
///      <see cref="UseKeywordHints"/>.
///
/// When nothing fires — or when both keyword lists match — the entry goes to both places,
/// which is the documented default.
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
    /// When false, only explicit markers route an entry and everything else goes to both places.
    /// Turn this off to get strictly predictable behaviour.
    /// </summary>
    public bool UseKeywordHints { get; set; } = true;

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
    /// Classifies by keyword. Used only when no explicit marker was given.
    /// A tie — no keywords, or keywords from both lists — means both.
    /// </summary>
    public EntryTarget ClassifyByKeyword(string text)
    {
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

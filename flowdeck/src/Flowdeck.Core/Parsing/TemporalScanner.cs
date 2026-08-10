using System.Globalization;
using System.Text.RegularExpressions;
using Flowdeck.Core.Models;

namespace Flowdeck.Core.Parsing;

/// <summary>
/// Turns raw Korean/English input into a flat list of recognised tokens.
/// Every regex contributes candidates; <see cref="Resolve"/> then keeps the
/// longest non-overlapping set, which is what makes "8월 12일" win over "12일"
/// and "1시간 전 알림" win over "1시간" without any ordering ceremony.
/// </summary>
internal static class TemporalScanner
{
    private const RegexOptions Opts = RegexOptions.CultureInvariant | RegexOptions.IgnoreCase;

    // ---- dates -------------------------------------------------------------
    private static readonly Regex IsoDate = new(@"(?<!\d)(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?!\d)", Opts);
    private static readonly Regex MonthDay = new(@"(?<!\d)(\d{1,2})\s*월\s*(\d{1,2})\s*일", Opts);
    private static readonly Regex SlashMonthDay = new(@"(?<!\d)(\d{1,2})/(\d{1,2})(?!\d)", Opts);
    private static readonly Regex RelDayWord = new(@"그저께|엊그제|어제|오늘|내일|모레|글피|today|tomorrow", Opts);
    private static readonly Regex RelWeekday = new(@"(이번\s*주|금주|다음\s*주|담주|차주|다다음\s*주|지난\s*주|저번\s*주)?\s*([월화수목금토일])\s*요일", Opts);
    private static readonly Regex RelWeekend = new(@"(이번|다음|담|다다음)?\s*주말", Opts);
    private static readonly Regex RelMonthDay = new(@"(이번\s*달|다음\s*달|담달|다다음\s*달)\s*(\d{1,2})\s*일", Opts);
    private static readonly Regex NthDay = new(@"(?<!\d)(\d{1,2})\s*일(?!\s*(?:후|뒤|간|동안|째|차|마다))", Opts);
    private static readonly Regex DaysLater = new(@"(?<!\d)(\d{1,3})\s*일\s*(?:후|뒤)", Opts);
    private static readonly Regex WeeksLater = new(@"(?<!\d)(\d{1,2})\s*주\s*(?:후|뒤)", Opts);
    private static readonly Regex WeekOnly = new(@"(다음\s*주|담주|차주)(?!\s*[월화수목금토일]\s*요일)", Opts);

    // ---- times -------------------------------------------------------------
    // The minute group stops short of "30분간" so that a trailing duration stays a duration:
    // "오후 2시 30분간 통화" is two o'clock for half an hour, not half past two.
    private static readonly Regex MeridiemTime =
        new(@"(오전|오후|아침|저녁|밤|새벽|점심)\s*(\d{1,2})\s*시(?:\s*(\d{1,2})\s*분(?!\s*(?:간|동안))|\s*(반))?", Opts);
    private static readonly Regex ClockTime = new(@"(?<!\d)(\d{1,2}):(\d{2})(?!\d)", Opts);
    private static readonly Regex BareTime =
        new(@"(?<!\d)(\d{1,2})\s*시(?!간)(?:\s*(\d{1,2})\s*분(?!\s*(?:간|동안))|\s*(반))?", Opts);
    private static readonly Regex DayPartOnly = new(@"(정오|자정|새벽|아침|점심|저녁|밤)(?=\s|$|\d)", Opts);
    private static readonly Regex EnglishTime = new(@"(?<![\w.])(\d{1,2})(?::(\d{2}))?\s*(am|pm)(?![\w])", Opts);

    // ---- everything else ---------------------------------------------------
    private static readonly Regex DurationHm = new(@"(?<!\d)(\d{1,2})\s*시간(?:\s*(\d{1,2})\s*분)?", Opts);
    private static readonly Regex DurationM = new(@"(?<!\d)(\d{1,3})\s*분\s*(?:간|동안)", Opts);
    private static readonly Regex Reminder =
        new(@"(?<!\d)(\d{1,3})\s*(분|시간)\s*전(?:\s*(?:에\s*)?(?:알림|리마인더|알려줘|알려))?", Opts);
    private static readonly Regex RecurEveryN = new(@"(?<!\d)(\d{1,2})\s*(일|주|개월|달|년)\s*마다", Opts);
    private static readonly Regex RecurWord = new(@"매일|평일마다|평일|주말마다|격주|격월|매주|매달|매월|매년|매해", Opts);
    private static readonly Regex PriorityTag = new(@"(?<![\w!])!p?([1-4])(?!\d)", Opts);
    private static readonly Regex PriorityWord = new(@"긴급|매우\s*중요|중요", Opts);
    private static readonly Regex HashTag = new(@"(?<![\w])#([^\s#]+)", Opts);

    // An address is taken as written, with or without the @ that introduces it. A bare
    // "@이름" is only a handle until the address book says otherwise, so it is matched
    // separately and may yet be handed back to the title.
    //
    // Both shapes come from ContactBook, which the address-book editor validates against.
    // Sharing the definition is what stops the editor accepting a name this never finds.
    private static readonly Regex AttendeeAddress =
        new(@"(?<![\w.@])@?(" + ContactBook.AddressPattern + @")(?![\w@])", Opts);

    private static readonly Regex AttendeeHandle =
        new(@"(?<![\w@])@(" + ContactBook.NamePattern + @")(?![\w@.])", Opts);
    private static readonly Regex RangeMark = new(@"[~～∼]|(?<=\d)\s*-\s*(?=\d{1,2}\s*(?:시|:))", Opts);

    /// <summary>Particles that trail a date or time and belong to it, not to the title.</summary>
    private static readonly Regex TrailingParticle =
        new(@"\G\s*(?:에는|에|엔|부터|까지|쯤|경|께|정도|사이)(?=\s|$)", Opts);

    private const int RankReminder = 100;
    private const int RankDuration = 90;
    private const int RankRecurrence = 85;
    private const int RankDate = 80;
    private const int RankTime = 70;
    private const int RankTag = 60;
    private const int RankAttendee = 58;
    private const int RankPriority = 55;
    private const int RankNoise = 5;

    public static List<Token> Scan(string text, DateTime now, DayOfWeek firstDayOfWeek = DayOfWeek.Monday)
    {
        var candidates = new List<Token>();

        AddReminders(text, candidates);
        AddDurations(text, candidates);
        AddRecurrences(text, candidates);
        AddDates(text, now, firstDayOfWeek, candidates);
        AddTimes(text, candidates);
        AddTags(text, candidates);
        AddAttendees(text, candidates);
        AddPriorities(text, candidates);

        foreach (Match m in RangeMark.Matches(text))
        {
            candidates.Add(new Token { Kind = TokenKind.Noise, Start = m.Index, Length = m.Length, Rank = RankNoise });
        }

        return Resolve(text, candidates);
    }

    /// <summary>
    /// Keeps the best non-overlapping subset: earliest match wins, then longest,
    /// then highest rank. Accepted date/time tokens absorb their trailing particle.
    /// </summary>
    private static List<Token> Resolve(string text, List<Token> candidates)
    {
        candidates.Sort((a, b) =>
        {
            var byStart = a.Start.CompareTo(b.Start);
            if (byStart != 0) return byStart;
            var byLength = b.Length.CompareTo(a.Length);
            if (byLength != 0) return byLength;
            return b.Rank.CompareTo(a.Rank);
        });

        var accepted = new List<Token>();
        var lastEnd = 0;

        foreach (var token in candidates)
        {
            if (token.Start < lastEnd) continue;

            if (token.Kind is TokenKind.Date or TokenKind.Time)
            {
                var particle = TrailingParticle.Match(text, token.End);
                if (particle.Success && particle.Index == token.End)
                {
                    token.Length += particle.Length;
                }
            }

            accepted.Add(token);
            lastEnd = token.End;
        }

        return accepted;
    }

    // ------------------------------------------------------------------------
    private static void AddReminders(string text, List<Token> into)
    {
        foreach (Match m in Reminder.Matches(text))
        {
            var value = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var minutes = m.Groups[2].Value == "시간" ? value * 60 : value;
            into.Add(new Token
            {
                Kind = TokenKind.Reminder,
                Start = m.Index,
                Length = m.Length,
                Rank = RankReminder,
                ReminderMinutes = minutes,
            });
        }
    }

    private static void AddDurations(string text, List<Token> into)
    {
        foreach (Match m in DurationHm.Matches(text))
        {
            var hours = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var minutes = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : 0;
            into.Add(new Token
            {
                Kind = TokenKind.Duration,
                Start = m.Index,
                Length = m.Length,
                Rank = RankDuration,
                Duration = new TimeSpan(hours, minutes, 0),
            });
        }

        foreach (Match m in DurationM.Matches(text))
        {
            var minutes = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            into.Add(new Token
            {
                Kind = TokenKind.Duration,
                Start = m.Index,
                Length = m.Length,
                Rank = RankDuration,
                Duration = TimeSpan.FromMinutes(minutes),
            });
        }
    }

    private static void AddRecurrences(string text, List<Token> into)
    {
        foreach (Match m in RecurEveryN.Matches(text))
        {
            var n = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var kind = m.Groups[2].Value switch
            {
                "일" => RecurrenceKind.Daily,
                "주" => RecurrenceKind.Weekly,
                "년" => RecurrenceKind.Yearly,
                _ => RecurrenceKind.Monthly,
            };
            into.Add(new Token
            {
                Kind = TokenKind.Recurrence,
                Start = m.Index,
                Length = m.Length,
                Rank = RankRecurrence,
                Recurrence = new Recurrence { Kind = kind, Interval = Math.Max(1, n) },
            });
        }

        foreach (Match m in RecurWord.Matches(text))
        {
            var recurrence = m.Value switch
            {
                "매일" => new Recurrence { Kind = RecurrenceKind.Daily, Interval = 1 },
                "평일" or "평일마다" => new Recurrence { Kind = RecurrenceKind.Weekdays, Interval = 1 },
                "주말마다" => new Recurrence
                {
                    Kind = RecurrenceKind.Weekly,
                    Interval = 1,
                    DaysOfWeek = { DayOfWeek.Saturday, DayOfWeek.Sunday },
                },
                "격주" => new Recurrence { Kind = RecurrenceKind.Weekly, Interval = 2 },
                "격월" => new Recurrence { Kind = RecurrenceKind.Monthly, Interval = 2 },
                "매주" => new Recurrence { Kind = RecurrenceKind.Weekly, Interval = 1 },
                "매달" or "매월" => new Recurrence { Kind = RecurrenceKind.Monthly, Interval = 1 },
                _ => new Recurrence { Kind = RecurrenceKind.Yearly, Interval = 1 },
            };

            into.Add(new Token
            {
                Kind = TokenKind.Recurrence,
                Start = m.Index,
                Length = m.Length,
                Rank = RankRecurrence,
                Recurrence = recurrence,
            });
        }
    }

    private static void AddDates(string text, DateTime now, DayOfWeek firstDayOfWeek, List<Token> into)
    {
        var today = now.Date;

        foreach (Match m in IsoDate.Matches(text))
        {
            var y = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var mo = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            var d = int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
            if (TryMakeDate(y, mo, d, out var date)) AddDate(into, m, date);
        }

        foreach (Match m in MonthDay.Matches(text))
        {
            var mo = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var d = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            if (!TryMakeDate(today.Year, mo, d, out var date)) continue;
            if (date < today) TryMakeDate(today.Year + 1, mo, d, out date);
            AddDate(into, m, date);
        }

        foreach (Match m in SlashMonthDay.Matches(text))
        {
            var mo = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var d = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            if (!TryMakeDate(today.Year, mo, d, out var date)) continue;
            if (date < today) TryMakeDate(today.Year + 1, mo, d, out date);
            AddDate(into, m, date);
        }

        foreach (Match m in RelDayWord.Matches(text))
        {
            var offset = m.Value.ToLowerInvariant() switch
            {
                "그저께" or "엊그제" => -2,
                "어제" => -1,
                "오늘" or "today" => 0,
                "내일" or "tomorrow" => 1,
                "모레" => 2,
                _ => 3, // 글피
            };
            AddDate(into, m, today.AddDays(offset));
        }

        foreach (Match m in DaysLater.Matches(text))
        {
            var n = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            AddDate(into, m, today.AddDays(n));
        }

        foreach (Match m in WeeksLater.Matches(text))
        {
            var n = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            AddDate(into, m, today.AddDays(7 * n));
        }

        foreach (Match m in RelWeekday.Matches(text))
        {
            var target = KoreanWeekday(m.Groups[2].Value);
            var qualifier = Normalize(m.Groups[1].Value);
            DateTime date;

            if (string.IsNullOrEmpty(qualifier))
            {
                // Bare weekday: the next one, counting today.
                var delta = ((int)target - (int)today.DayOfWeek + 7) % 7;
                date = today.AddDays(delta);
            }
            else
            {
                var weekShift = qualifier switch
                {
                    "다음주" or "담주" or "차주" => 7,
                    "다다음주" => 14,
                    "지난주" or "저번주" => -7,
                    _ => 0, // 이번주, 금주
                };
                date = StartOfWeek(today, firstDayOfWeek)
                    .AddDays(weekShift + WeekIndex(target, firstDayOfWeek));
            }

            AddDate(into, m, date);
        }

        foreach (Match m in RelWeekend.Matches(text))
        {
            var qualifier = Normalize(m.Groups[1].Value);
            var weekShift = qualifier switch
            {
                "다음" or "담" => 7,
                "다다음" => 14,
                _ => 0,
            };
            AddDate(
                into,
                m,
                StartOfWeek(today, firstDayOfWeek)
                    .AddDays(weekShift + WeekIndex(DayOfWeek.Saturday, firstDayOfWeek)));
        }

        foreach (Match m in RelMonthDay.Matches(text))
        {
            var qualifier = Normalize(m.Groups[1].Value);
            var monthShift = qualifier switch
            {
                "다음달" or "담달" => 1,
                "다다음달" => 2,
                _ => 0,
            };
            var d = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            var anchor = today.AddMonths(monthShift);
            if (TryMakeDate(anchor.Year, anchor.Month, d, out var date)) AddDate(into, m, date);
        }

        foreach (Match m in WeekOnly.Matches(text))
        {
            AddDate(into, m, StartOfWeek(today, firstDayOfWeek).AddDays(7));
        }

        foreach (Match m in NthDay.Matches(text))
        {
            var d = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var anchor = d < today.Day ? today.AddMonths(1) : today;
            if (TryMakeDate(anchor.Year, anchor.Month, d, out var date)) AddDate(into, m, date);
        }
    }

    private static void AddTimes(string text, List<Token> into)
    {
        foreach (Match m in MeridiemTime.Matches(text))
        {
            var hour = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            if (hour > 24) continue;
            var minute = m.Groups[3].Success
                ? int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture)
                : m.Groups[4].Success ? 30 : 0;
            if (minute > 59) continue;

            hour = ApplyMeridiem(m.Groups[1].Value, hour);
            AddTime(into, m, hour, minute, explicitMeridiem: true);
        }

        foreach (Match m in ClockTime.Matches(text))
        {
            var hour = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            var minute = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            if (hour > 23 || minute > 59) continue;
            AddTime(into, m, hour, minute, explicitMeridiem: hour >= 13);
        }

        foreach (Match m in BareTime.Matches(text))
        {
            var hour = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            if (hour > 23) continue;
            var minute = m.Groups[2].Success
                ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture)
                : m.Groups[3].Success ? 30 : 0;
            if (minute > 59) continue;
            AddTime(into, m, hour, minute, explicitMeridiem: hour >= 13);
        }

        foreach (Match m in EnglishTime.Matches(text))
        {
            var hour = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            if (hour > 12) continue;
            var minute = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) : 0;
            if (minute > 59) continue;
            var isPm = m.Groups[3].Value.Equals("pm", StringComparison.OrdinalIgnoreCase);
            hour = isPm ? (hour == 12 ? 12 : hour + 12) : (hour == 12 ? 0 : hour);
            AddTime(into, m, hour, minute, explicitMeridiem: true);
        }

        foreach (Match m in DayPartOnly.Matches(text))
        {
            var hour = m.Groups[1].Value switch
            {
                "자정" => 0,
                "새벽" => 6,
                "아침" => 8,
                "정오" or "점심" => 12,
                "저녁" => 18,
                _ => 21, // 밤
            };
            AddTime(into, m, hour, 0, explicitMeridiem: true);
        }
    }

    private static void AddTags(string text, List<Token> into)
    {
        foreach (Match m in HashTag.Matches(text))
        {
            into.Add(new Token
            {
                Kind = TokenKind.Tag,
                Start = m.Index,
                Length = m.Length,
                Rank = RankTag,
                Text = m.Groups[1].Value,
            });
        }
    }

    /// <summary>
    /// Picks out who a meeting is with. An address wins over a bare handle wherever both
    /// could match, because "@hong@corp.com" is longer than the "@hong" inside it and the
    /// longest candidate is the one that survives.
    /// </summary>
    private static void AddAttendees(string text, List<Token> into)
    {
        foreach (Match m in AttendeeAddress.Matches(text))
        {
            into.Add(new Token
            {
                Kind = TokenKind.Attendee,
                Start = m.Index,
                Length = m.Length,
                Rank = RankAttendee,
                Text = m.Groups[1].Value,
            });
        }

        foreach (Match m in AttendeeHandle.Matches(text))
        {
            into.Add(new Token
            {
                Kind = TokenKind.Attendee,
                Start = m.Index,
                Length = m.Length,
                Rank = RankAttendee,
                Text = m.Groups[1].Value,
            });
        }
    }

    private static void AddPriorities(string text, List<Token> into)
    {
        foreach (Match m in PriorityTag.Matches(text))
        {
            var level = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            into.Add(new Token
            {
                Kind = TokenKind.Priority,
                Start = m.Index,
                Length = m.Length,
                Rank = RankPriority,
                Priority = level switch
                {
                    1 => Priority.Urgent,
                    2 => Priority.High,
                    3 => Priority.Normal,
                    _ => Priority.Low,
                },
            });
        }

        // Words like "긴급" set the priority but stay in the title, because removing
        // them would quietly change what the entry says.
        foreach (Match m in PriorityWord.Matches(text))
        {
            var priority = m.Value.StartsWith("긴급", StringComparison.Ordinal) ? Priority.Urgent : Priority.High;
            into.Add(new Token
            {
                Kind = TokenKind.Priority,
                Start = m.Index,
                Length = m.Length,
                Rank = RankPriority,
                Strip = false,
                Priority = priority,
            });
        }
    }

    // ------------------------------------------------------------------------
    private static void AddDate(List<Token> into, Match m, DateTime date) =>
        into.Add(new Token
        {
            Kind = TokenKind.Date,
            Start = m.Index,
            Length = m.Length,
            Rank = RankDate,
            Date = date.Date,
        });

    private static void AddTime(List<Token> into, Match m, int hour, int minute, bool explicitMeridiem) =>
        into.Add(new Token
        {
            Kind = TokenKind.Time,
            Start = m.Index,
            Length = m.Length,
            Rank = RankTime,
            Time = new TimeFragment { Hour = hour, Minute = minute, ExplicitMeridiem = explicitMeridiem },
        });

    private static int ApplyMeridiem(string marker, int hour)
    {
        var isMorning = marker is "오전" or "아침" or "새벽";
        if (isMorning) return hour == 12 ? 0 : hour;
        return hour < 12 ? hour + 12 : 12; // 오후/저녁/밤/점심
    }

    private static bool TryMakeDate(int year, int month, int day, out DateTime date)
    {
        date = default;
        if (year is < 1 or > 9999 || month is < 1 or > 12 || day < 1) return false;
        if (day > DateTime.DaysInMonth(year, month)) return false;
        date = new DateTime(year, month, day);
        return true;
    }

    private static string Normalize(string s) => s.Replace(" ", string.Empty).Replace("\t", string.Empty);

    private static DayOfWeek KoreanWeekday(string s) => s switch
    {
        "월" => DayOfWeek.Monday,
        "화" => DayOfWeek.Tuesday,
        "수" => DayOfWeek.Wednesday,
        "목" => DayOfWeek.Thursday,
        "금" => DayOfWeek.Friday,
        "토" => DayOfWeek.Saturday,
        _ => DayOfWeek.Sunday,
    };

    private static DateTime StartOfWeek(DateTime date, DayOfWeek firstDayOfWeek) =>
        date.AddDays(-WeekIndex(date.DayOfWeek, firstDayOfWeek));

    /// <summary>Position of <paramref name="day"/> within a week that opens on <paramref name="firstDayOfWeek"/>.</summary>
    private static int WeekIndex(DayOfWeek day, DayOfWeek firstDayOfWeek) =>
        ((int)day - (int)firstDayOfWeek + 7) % 7;
}

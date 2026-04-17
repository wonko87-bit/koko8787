package com.taskbridge.parser

import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.ZonedDateTime
import java.time.format.DateTimeFormatter

private val KST = ZoneId.of("Asia/Seoul")

data class ParsedDateTime(
    val start: ZonedDateTime?,   // null이면 종일
    val end: ZonedDateTime?,
    val startDate: LocalDate,
    val endDate: LocalDate,
    val hasTime: Boolean,
)

object DateTimeParser {

    // ---- 정규식 ----
    private val MONTH_DAY   = Regex("""(\d{1,2})월\s*(\d{1,2})일""")
    private val DAY_ONLY    = Regex("""(?<!\d)(\d{1,2})일(?!\w)""")
    private val AMPM_TIME   = Regex("""(오전|오후)\s*(\d{1,2})시(?:\s*(\d{1,2})분)?""")
    private val BARE_TIME   = Regex("""(?<!\d)(\d{1,2})시(?:\s*(\d{1,2})분)?""")
    private val CLOCK_TIME  = Regex("""(\d{1,2}):(\d{2})""")
    private val META_SYMBOL = Regex("""[#@!]\S*""")
    private val SLASH_PREFIX = Regex("""^/[ct]\s*""")

    // ---- strip: 날짜·시간·메타 기호 제거 ----
    fun strip(text: String): String {
        var s = SLASH_PREFIX.replace(text, "")
        // ~ 이후 끝 시각 제거
        s = Regex("""\s*~\s*(?:(?:\d{1,2}월\s*\d{1,2}일\s*)?(?:오전|오후)?\s*\d{1,2}시(?:\s*\d{1,2}분)?|\d{1,2}:\d{2})""").replace(s, "")
        s = AMPM_TIME.replace(s, "")
        s = CLOCK_TIME.replace(s, "")
        s = BARE_TIME.replace(s, "")
        s = Regex("""\d{4}-\d{2}-\d{2}""").replace(s, "")
        s = MONTH_DAY.replace(s, "")
        s = DAY_ONLY.replace(s, "")
        s = Regex("""오늘|내일|모레""").replace(s, "")
        s = META_SYMBOL.replace(s, "")
        return s.replace(Regex("""\s+"""), " ").trim()
    }

    // ---- 단일 datetime 파싱 ----
    private fun parseSingle(text: String, defaultDate: LocalDate? = null): Triple<LocalDate, Int?, Int> {
        val today = LocalDate.now(KST)
        var base  = defaultDate ?: today

        // 날짜
        when {
            "내일" in text -> base = today.plusDays(1)
            "모레" in text -> base = today.plusDays(2)
            "오늘" in text -> base = today
            else -> {
                MONTH_DAY.find(text)?.let {
                    runCatching { base = today.withMonth(it.groupValues[1].toInt()).withDayOfMonth(it.groupValues[2].toInt()) }
                } ?: DAY_ONLY.find(text)?.let {
                    runCatching { base = today.withDayOfMonth(it.groupValues[1].toInt()) }
                } ?: Regex("""(\d{4})-(\d{2})-(\d{2})""").find(text)?.let {
                    runCatching { base = LocalDate.of(it.groupValues[1].toInt(), it.groupValues[2].toInt(), it.groupValues[3].toInt()) }
                }
            }
        }

        // 시각
        var hour: Int? = null
        var minute = 0

        AMPM_TIME.find(text)?.let { m ->
            var h = m.groupValues[2].toInt()
            val ampm = m.groupValues[1]
            if (ampm == "오후" && h != 12) h += 12
            else if (ampm == "오전" && h == 12) h = 0
            hour   = h
            minute = m.groupValues[3].toIntOrNull() ?: 0
        } ?: CLOCK_TIME.find(text)?.let { m ->
            hour   = m.groupValues[1].toInt()
            minute = m.groupValues[2].toInt()
        } ?: BARE_TIME.find(text)?.let { m ->
            hour   = m.groupValues[1].toInt()
            minute = m.groupValues[2].toIntOrNull() ?: 0
        }

        return Triple(base, hour, minute)
    }

    private fun hasDateExpr(text: String) =
        Regex("""오늘|내일|모레|\d{1,2}월\s*\d{1,2}일|(?<!\d)\d{1,2}일(?!\w)""").containsMatchIn(text)

    // ---- 메인 파싱 (범위 포함) ----
    fun parse(text: String): ParsedDateTime? {
        if ("~" in text) {
            val (startRaw, _, endRaw) = text.partition { it == '~' }.let {
                val idx = text.indexOf('~')
                Triple(text.substring(0, idx).trim(), '~', text.substring(idx + 1).trim())
            }
            val (sd, sh, sm) = parseSingle(startRaw)
            val (ed, eh, em) = parseSingle(endRaw, defaultDate = sd)

            // 오후 맥락 전파
            var finalEh = eh
            if (sh != null && eh != null && !Regex("""오전|오후""").containsMatchIn(endRaw)) {
                if (sh >= 12 && eh < sh && eh < 12) finalEh = eh + 12
            }

            if (sh != null && finalEh != null) {
                val s = ZonedDateTime.of(LocalDateTime.of(sd, java.time.LocalTime.of(sh, sm)), KST)
                val e = ZonedDateTime.of(LocalDateTime.of(ed, java.time.LocalTime.of(finalEh, em)), KST)
                return ParsedDateTime(s, e, sd, ed, true)
            }
            if (hasDateExpr(startRaw) || hasDateExpr(endRaw)) {
                return ParsedDateTime(null, null, sd, ed, false)
            }
            return null
        }

        // 단일
        val (base, hour, minute) = parseSingle(text)
        if (hour != null) {
            val s = ZonedDateTime.of(LocalDateTime.of(base, java.time.LocalTime.of(hour, minute)), KST)
            return ParsedDateTime(s, s.plusHours(1), base, base, true)
        }
        if (hasDateExpr(text)) {
            return ParsedDateTime(null, null, base, base, false)
        }
        return null
    }

    // Todoist용: 시간 있으면 (RFC3339 UTC, true), 없으면 (YYYY-MM-DD, false)
    fun extractDueInfo(text: String): Pair<String, Boolean>? {
        val p = parse(text) ?: return null
        return if (p.hasTime && p.start != null) {
            val utc = p.start.withZoneSameInstant(ZoneId.of("UTC"))
            Pair(utc.format(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss'Z'")), true)
        } else {
            Pair(p.startDate.toString(), false)
        }
    }
}

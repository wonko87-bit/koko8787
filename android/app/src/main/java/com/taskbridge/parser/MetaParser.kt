package com.taskbridge.parser

data class TodoistMeta(
    val quickAddText: String,  // Todoist Quick Add에 보낼 텍스트 (!r 제거, 반복 영어 변환)
    val reminderMinutes: Int?, // !r30 → 30
)

object MetaParser {

    private val REMINDER_RE = Regex("""!r(\d+)\b""", RegexOption.IGNORE_CASE)

    // 한국어 반복 키워드 → Todoist Quick Add가 이해하는 영어로 변환
    private val RECURRENCE_MAP = listOf(
        "매격주" to "every 2 weeks",
        "격주"   to "every 2 weeks",
        "매주"   to "every week",
        "매달"   to "every month",
        "매월"   to "every month",
        "매일"   to "every day",
    )

    fun extractReminderMinutes(text: String): Int? =
        REMINDER_RE.find(text)?.groupValues?.get(1)?.toIntOrNull()

    fun parse(text: String): TodoistMeta {
        var s = text

        // !r30 제거
        val remMatch = REMINDER_RE.find(s)
        val reminderMinutes = remMatch?.groupValues?.get(1)?.toIntOrNull()
        if (remMatch != null) s = s.removeRange(remMatch.range).trim()

        // 반복 키워드 영어로 변환
        for ((korean, english) in RECURRENCE_MAP) {
            if (korean in s) { s = s.replace(korean, english); break }
        }

        return TodoistMeta(
            quickAddText   = s.replace(Regex("""\s+"""), " ").trim(),
            reminderMinutes = reminderMinutes,
        )
    }
}

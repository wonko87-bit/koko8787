package com.taskbridge.parser

data class TodoistMeta(
    val content: String,
    val project: String?,
    val labels: List<String>,
    val priority: Int?,   // Todoist 내부값: 4=긴급, 3=높음, 2=보통, 1=낮음
)

object MetaParser {

    private val PRIORITY_RE = Regex("""!p?([1-4])\b""", RegexOption.IGNORE_CASE)
    private val PROJECT_RE  = Regex("""#(\S+)""")
    private val LABEL_RE    = Regex("""@(\S+)""")

    fun parse(text: String): TodoistMeta {
        var s = text

        // 우선순위
        val priMatch = PRIORITY_RE.find(s)
        val priority = priMatch?.groupValues?.get(1)?.toIntOrNull()?.let { 5 - it }
        if (priMatch != null) s = s.removeRange(priMatch.range).trim()

        // 프로젝트
        val projMatch = PROJECT_RE.find(s)
        val project = projMatch?.groupValues?.get(1)
        if (projMatch != null) s = s.removeRange(projMatch.range).trim()

        // 레이블
        val labels = LABEL_RE.findAll(s).map { it.groupValues[1] }.toList()
        s = LABEL_RE.replace(s, "").trim()

        // 날짜/시간 제거
        val clean = DateTimeParser.strip(s).ifBlank { s }

        return TodoistMeta(
            content  = clean,
            project  = project,
            labels   = labels,
            priority = priority,
        )
    }
}

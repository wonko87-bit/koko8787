package com.taskbridge.parser

enum class Target { BOTH, CALENDAR, TODOIST }

data class ParsedInput(val target: Target, val text: String)

object InputParser {
    fun parse(raw: String): ParsedInput {
        val s = raw.trim()
        return when {
            s.startsWith("/c ") -> ParsedInput(Target.CALENDAR, s.removePrefix("/c ").trim())
            s == "/c"           -> ParsedInput(Target.CALENDAR, "")
            s.startsWith("/t ") -> ParsedInput(Target.TODOIST,  s.removePrefix("/t ").trim())
            s == "/t"           -> ParsedInput(Target.TODOIST,  "")
            else                -> ParsedInput(Target.BOTH, s)
        }
    }
}

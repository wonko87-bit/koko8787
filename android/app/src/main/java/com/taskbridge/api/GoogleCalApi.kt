package com.taskbridge.api

import com.taskbridge.parser.DateTimeParser
import com.taskbridge.parser.MetaParser
import com.taskbridge.parser.RecurrenceType
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.OkHttpClient
import okhttp3.Request
import okhttp3.RequestBody.Companion.toRequestBody
import org.json.JSONArray
import org.json.JSONObject
import java.io.IOException
import java.time.LocalDate
import java.time.ZoneId
import java.time.format.DateTimeFormatter

object GoogleCalApi {

    private const val BASE = "https://www.googleapis.com/calendar/v3/calendars/primary/events"
    private val JSON_MT  = "application/json".toMediaType()
    private val http     = OkHttpClient()
    private val DT_FMT   = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ssxxx")
    private val DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd")

    fun createEvent(token: String, rawText: String): JSONObject {
        val dt         = DateTimeParser.parse(rawText)
        val recurrence = DateTimeParser.extractRecurrence(rawText)
        val reminder   = MetaParser.extractReminderMinutes(rawText)
        val summary    = DateTimeParser.strip(rawText).ifBlank { rawText }

        val payload = JSONObject().put("summary", summary)

        // 날짜/시간
        if (dt != null && dt.hasTime && dt.start != null && dt.end != null) {
            payload.put("start", JSONObject().put("dateTime", dt.start.format(DT_FMT)))
            payload.put("end",   JSONObject().put("dateTime", dt.end.format(DT_FMT)))
        } else if (dt != null) {
            payload.put("start", JSONObject().put("date", dt.startDate.format(DATE_FMT)))
            payload.put("end",   JSONObject().put("date", dt.endDate.plusDays(1).format(DATE_FMT)))
        } else {
            val today = LocalDate.now(ZoneId.of("Asia/Seoul")).format(DATE_FMT)
            payload.put("start", JSONObject().put("date", today))
            payload.put("end",   JSONObject().put("date", today))
        }

        // 반복
        if (recurrence != null) {
            val rrule = when (recurrence) {
                RecurrenceType.DAILY    -> "RRULE:FREQ=DAILY"
                RecurrenceType.WEEKLY   -> "RRULE:FREQ=WEEKLY"
                RecurrenceType.BIWEEKLY -> "RRULE:FREQ=WEEKLY;INTERVAL=2"
                RecurrenceType.MONTHLY  -> "RRULE:FREQ=MONTHLY"
            }
            payload.put("recurrence", JSONArray().put(rrule))
        }

        // 리마인더
        if (reminder != null) {
            payload.put("reminders", JSONObject()
                .put("useDefault", false)
                .put("overrides", JSONArray().put(
                    JSONObject().put("method", "popup").put("minutes", reminder)
                ))
            )
        }

        val resp = http.newCall(
            Request.Builder()
                .url(BASE)
                .header("Authorization", "Bearer $token")
                .post(payload.toString().toRequestBody(JSON_MT))
                .build()
        ).execute()

        if (!resp.isSuccessful) throw IOException("Google Calendar failed: ${resp.code} ${resp.body?.string()}")
        return JSONObject(resp.body!!.string())
    }
}

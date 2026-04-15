package com.taskbridge.api

import com.taskbridge.parser.DateTimeParser
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.OkHttpClient
import okhttp3.Request
import okhttp3.RequestBody.Companion.toRequestBody
import org.json.JSONObject
import java.io.IOException
import java.time.LocalDate
import java.time.ZoneId
import java.time.format.DateTimeFormatter

object GoogleCalApi {

    private const val BASE = "https://www.googleapis.com/calendar/v3/calendars/primary/events"
    private val JSON_MT = "application/json".toMediaType()
    private val http = OkHttpClient()
    private val DT_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ssxxx")
    private val DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd")

    fun createEvent(token: String, rawText: String): JSONObject {
        val dt      = DateTimeParser.parse(rawText)
        val summary = DateTimeParser.strip(rawText).ifBlank { rawText }

        val payload = JSONObject().put("summary", summary)

        if (dt != null && dt.hasTime && dt.start != null && dt.end != null) {
            payload.put("start", JSONObject().put("dateTime", dt.start.format(DT_FMT)))
            payload.put("end",   JSONObject().put("dateTime", dt.end.format(DT_FMT)))
        } else if (dt != null) {
            payload.put("start", JSONObject().put("date", dt.startDate.format(DATE_FMT)))
            payload.put("end",   JSONObject().put("date", dt.endDate.plusDays(1).format(DATE_FMT)))
        } else {
            // 날짜 없음 → 오늘 종일
            val today = LocalDate.now(ZoneId.of("Asia/Seoul")).format(DATE_FMT)
            payload.put("start", JSONObject().put("date", today))
            payload.put("end",   JSONObject().put("date", today))
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

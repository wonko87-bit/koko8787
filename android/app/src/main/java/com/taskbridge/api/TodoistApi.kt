package com.taskbridge.api

import com.taskbridge.parser.MetaParser
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.OkHttpClient
import okhttp3.Request
import okhttp3.RequestBody.Companion.toRequestBody
import org.json.JSONObject
import java.io.IOException

object TodoistApi {

    private const val BASE = "https://api.todoist.com/api/v1"
    private val JSON_MT = "application/json".toMediaType()
    private val http = OkHttpClient()

    private fun post(token: String, path: String, body: JSONObject): JSONObject {
        val resp = http.newCall(
            Request.Builder()
                .url("$BASE$path")
                .header("Authorization", "Bearer $token")
                .post(body.toString().toRequestBody(JSON_MT))
                .build()
        ).execute()
        if (!resp.isSuccessful) throw IOException("Todoist POST $path failed: ${resp.code} ${resp.body?.string()}")
        return JSONObject(resp.body!!.string())
    }

    fun createTask(token: String, rawContent: String): JSONObject {
        val meta = MetaParser.parse(rawContent)

        // Quick Add: Todoist가 날짜·시간·반복·프로젝트·레이블·우선순위를 직접 파싱
        val task = post(token, "/tasks/quick", JSONObject().put("text", meta.quickAddText))

        // 리마인더 (!r30 → due 기준 30분 전 알림)
        meta.reminderMinutes?.let { mins ->
            runCatching {
                post(token, "/reminders", JSONObject()
                    .put("task_id",      task.getString("id"))
                    .put("type",         "relative")
                    .put("minute_offset", mins)
                )
            }
        }

        return task
    }
}

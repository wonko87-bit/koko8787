package com.taskbridge.api

import com.taskbridge.parser.MetaParser
import com.taskbridge.parser.DateTimeParser
import okhttp3.MediaType.Companion.toMediaType
import okhttp3.OkHttpClient
import okhttp3.Request
import okhttp3.RequestBody.Companion.toRequestBody
import org.json.JSONArray
import org.json.JSONObject
import java.io.IOException

object TodoistApi {

    private const val BASE = "https://api.todoist.com/api/v1"
    private val JSON_MT = "application/json".toMediaType()
    private val http = OkHttpClient()

    private fun get(token: String, path: String): JSONArray {
        val resp = http.newCall(
            Request.Builder()
                .url("$BASE$path")
                .header("Authorization", "Bearer $token")
                .get().build()
        ).execute()
        if (!resp.isSuccessful) throw IOException("Todoist GET $path failed: ${resp.code}")
        return JSONArray(resp.body!!.string())
    }

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

    // ---- 프로젝트 조회 or 생성 ----
    private fun getOrCreateProjectId(token: String, name: String): String? {
        val projects = get(token, "/projects")
        for (i in 0 until projects.length()) {
            val p = projects.getJSONObject(i)
            if (p.getString("name").equals(name, ignoreCase = true)) return p.getString("id")
        }
        // 없으면 생성
        return runCatching {
            post(token, "/projects", JSONObject().put("name", name)).getString("id")
        }.getOrNull()
    }

    // ---- 태스크 생성 ----
    fun createTask(token: String, rawContent: String): JSONObject {
        val meta = MetaParser.parse(rawContent)
        val payload = JSONObject().put("content", meta.content.ifBlank { rawContent })

        // 날짜
        DateTimeParser.extractDueDate(rawContent)?.let { payload.put("due_date", it) }

        // 우선순위
        meta.priority?.let { payload.put("priority", it) }

        // 레이블
        if (meta.labels.isNotEmpty()) {
            payload.put("labels", JSONArray(meta.labels))
        }

        // 프로젝트
        meta.project?.let { name ->
            getOrCreateProjectId(token, name)?.let { payload.put("project_id", it) }
        }

        return post(token, "/tasks", payload)
    }
}

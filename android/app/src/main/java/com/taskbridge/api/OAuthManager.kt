package com.taskbridge.api

import android.content.Context
import android.net.Uri
import com.taskbridge.storage.TokenStore
import okhttp3.FormBody
import okhttp3.OkHttpClient
import okhttp3.Request
import org.json.JSONObject
import java.io.IOException

object OAuthManager {

    // GitHub Secrets → build.gradle → BuildConfig 를 통해 자동 주입됨
    val GOOGLE_CLIENT_ID      get() = com.taskbridge.BuildConfig.GOOGLE_CLIENT_ID
    val GOOGLE_CLIENT_SECRET  get() = com.taskbridge.BuildConfig.GOOGLE_CLIENT_SECRET
    val TODOIST_CLIENT_ID     get() = com.taskbridge.BuildConfig.TODOIST_CLIENT_ID
    val TODOIST_CLIENT_SECRET get() = com.taskbridge.BuildConfig.TODOIST_CLIENT_SECRET

    const val GOOGLE_REDIRECT_URI  = "https://koko8787-production.up.railway.app/auth/mobile/google/callback"
    const val TODOIST_REDIRECT_URI = "https://koko8787-production.up.railway.app/auth/mobile/todoist/callback"

    private val GOOGLE_SCOPES = "https://www.googleapis.com/auth/calendar.events"

    private val http = OkHttpClient()

    // ---- Google 인증 URL 생성 ----
    fun googleAuthUrl(state: String): String =
        Uri.parse("https://accounts.google.com/o/oauth2/v2/auth").buildUpon()
            .appendQueryParameter("client_id",     GOOGLE_CLIENT_ID)
            .appendQueryParameter("redirect_uri",  GOOGLE_REDIRECT_URI)
            .appendQueryParameter("response_type", "code")
            .appendQueryParameter("scope",         GOOGLE_SCOPES)
            .appendQueryParameter("access_type",   "offline")
            .appendQueryParameter("prompt",        "consent")
            .appendQueryParameter("state",         state)
            .build().toString()

    // ---- Todoist 인증 URL 생성 ----
    fun todoistAuthUrl(state: String): String =
        Uri.parse("https://todoist.com/oauth/authorize").buildUpon()
            .appendQueryParameter("client_id", TODOIST_CLIENT_ID)
            .appendQueryParameter("scope",     "task:add,data:read_write")
            .appendQueryParameter("state",     state)
            .build().toString()

    // ---- Google 코드 교환 ----
    fun exchangeGoogleCode(ctx: Context, code: String) {
        val body = FormBody.Builder()
            .add("code",          code)
            .add("client_id",     GOOGLE_CLIENT_ID)
            .add("client_secret", GOOGLE_CLIENT_SECRET)
            .add("redirect_uri",  GOOGLE_REDIRECT_URI)
            .add("grant_type",    "authorization_code")
            .build()

        val resp = http.newCall(
            Request.Builder().url("https://oauth2.googleapis.com/token").post(body).build()
        ).execute()

        if (!resp.isSuccessful) throw IOException("Google token exchange failed: ${resp.code}")

        val json = JSONObject(resp.body!!.string())
        TokenStore.saveGoogle(
            ctx,
            accessToken  = json.getString("access_token"),
            refreshToken = json.optString("refresh_token").ifBlank { null },
            expiresIn    = json.optLong("expires_in", 3600L),
        )
    }

    // ---- Google 토큰 갱신 ----
    fun refreshGoogleToken(ctx: Context): Boolean {
        val refresh = TokenStore.googleRefreshToken(ctx) ?: return false
        val body = FormBody.Builder()
            .add("client_id",     GOOGLE_CLIENT_ID)
            .add("client_secret", GOOGLE_CLIENT_SECRET)
            .add("refresh_token", refresh)
            .add("grant_type",    "refresh_token")
            .build()
        val resp = http.newCall(
            Request.Builder().url("https://oauth2.googleapis.com/token").post(body).build()
        ).execute()
        if (!resp.isSuccessful) return false
        val json = JSONObject(resp.body!!.string())
        TokenStore.saveGoogle(
            ctx,
            accessToken  = json.getString("access_token"),
            refreshToken = null,
            expiresIn    = json.optLong("expires_in", 3600L),
        )
        return true
    }

    // ---- Todoist 코드 교환 ----
    fun exchangeTodoistCode(ctx: Context, code: String) {
        val body = FormBody.Builder()
            .add("code",          code)
            .add("client_id",     TODOIST_CLIENT_ID)
            .add("client_secret", TODOIST_CLIENT_SECRET)
            .build()

        val resp = http.newCall(
            Request.Builder().url("https://todoist.com/oauth/access_token").post(body).build()
        ).execute()

        if (!resp.isSuccessful) throw IOException("Todoist token exchange failed: ${resp.code}")

        val json = JSONObject(resp.body!!.string())
        TokenStore.saveTodoist(ctx, json.getString("access_token"))
    }

    // ---- 유효한 Google 액세스 토큰 반환 (만료 시 자동 갱신) ----
    fun validGoogleToken(ctx: Context): String? {
        val expiry = TokenStore.googleExpiry(ctx)
        if (System.currentTimeMillis() > expiry - 60_000) {
            if (!refreshGoogleToken(ctx)) return null
        }
        return TokenStore.googleAccessToken(ctx)
    }
}

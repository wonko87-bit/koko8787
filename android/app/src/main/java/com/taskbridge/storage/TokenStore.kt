package com.taskbridge.storage

import android.content.Context
import androidx.security.crypto.EncryptedSharedPreferences
import androidx.security.crypto.MasterKey

/**
 * 암호화된 SharedPreferences에 OAuth 토큰을 저장/조회한다.
 * 앱 삭제 전까지 영구 보존.
 */
object TokenStore {

    private const val FILE = "taskbridge_tokens"
    private const val KEY_GOOGLE_ACCESS  = "google_access_token"
    private const val KEY_GOOGLE_REFRESH = "google_refresh_token"
    private const val KEY_GOOGLE_EXPIRY  = "google_token_expiry"
    private const val KEY_TODOIST        = "todoist_access_token"

    private fun prefs(ctx: Context) = EncryptedSharedPreferences.create(
        ctx,
        FILE,
        MasterKey.Builder(ctx).setKeyScheme(MasterKey.KeyScheme.AES256_GCM).build(),
        EncryptedSharedPreferences.PrefKeyEncryptionScheme.AES256_SIV,
        EncryptedSharedPreferences.PrefValueEncryptionScheme.AES256_GCM,
    )

    // ---- Google ----

    fun saveGoogle(ctx: Context, accessToken: String, refreshToken: String?, expiresIn: Long) {
        prefs(ctx).edit().apply {
            putString(KEY_GOOGLE_ACCESS, accessToken)
            if (refreshToken != null) putString(KEY_GOOGLE_REFRESH, refreshToken)
            putLong(KEY_GOOGLE_EXPIRY, System.currentTimeMillis() + expiresIn * 1000)
            apply()
        }
    }

    fun googleAccessToken(ctx: Context): String? = prefs(ctx).getString(KEY_GOOGLE_ACCESS, null)
    fun googleRefreshToken(ctx: Context): String? = prefs(ctx).getString(KEY_GOOGLE_REFRESH, null)
    fun googleExpiry(ctx: Context): Long = prefs(ctx).getLong(KEY_GOOGLE_EXPIRY, 0L)
    fun isGoogleConnected(ctx: Context): Boolean = googleAccessToken(ctx) != null

    fun clearGoogle(ctx: Context) {
        prefs(ctx).edit()
            .remove(KEY_GOOGLE_ACCESS).remove(KEY_GOOGLE_REFRESH).remove(KEY_GOOGLE_EXPIRY)
            .apply()
    }

    // ---- Todoist ----

    fun saveTodoist(ctx: Context, accessToken: String) {
        prefs(ctx).edit().putString(KEY_TODOIST, accessToken).apply()
    }

    fun todoistAccessToken(ctx: Context): String? = prefs(ctx).getString(KEY_TODOIST, null)
    fun isTodoistConnected(ctx: Context): Boolean = todoistAccessToken(ctx) != null

    fun clearTodoist(ctx: Context) {
        prefs(ctx).edit().remove(KEY_TODOIST).apply()
    }
}

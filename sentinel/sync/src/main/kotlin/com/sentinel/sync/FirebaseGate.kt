package com.sentinel.sync

import android.content.Context
import com.google.firebase.FirebaseApp
import com.google.firebase.FirebaseOptions
import com.google.firebase.firestore.FirebaseFirestore

/**
 * google-services.json 없이 값 주입만으로 Firebase를 수동 초기화한다.
 * 설정값(BuildConfig FB_*)이 비어 있으면 [firestore] == null → 동기화 자동 비활성.
 *
 * 실기 활성화: gradle.properties 또는 환경변수에
 *   SENTINEL_FB_APP_ID / SENTINEL_FB_API_KEY / SENTINEL_FB_PROJECT_ID
 * 를 채우면 된다. (Firebase 콘솔 → 프로젝트 설정)
 */
object FirebaseGate {

    private const val APP_NAME = "sentinel"

    val isConfigured: Boolean
        get() = BuildConfig.FB_APP_ID.isNotBlank() &&
            BuildConfig.FB_API_KEY.isNotBlank() &&
            BuildConfig.FB_PROJECT_ID.isNotBlank()

    @Volatile
    private var cached: FirebaseFirestore? = null

    /** 설정돼 있으면 Firestore 인스턴스, 아니면 null. */
    fun firestore(context: Context): FirebaseFirestore? {
        if (!isConfigured) return null
        cached?.let { return it }
        return synchronized(this) {
            cached ?: run {
                val app = existingApp() ?: FirebaseApp.initializeApp(
                    context.applicationContext,
                    FirebaseOptions.Builder()
                        .setApplicationId(BuildConfig.FB_APP_ID)
                        .setApiKey(BuildConfig.FB_API_KEY)
                        .setProjectId(BuildConfig.FB_PROJECT_ID)
                        .build(),
                    APP_NAME,
                )
                FirebaseFirestore.getInstance(app).also { cached = it }
            }
        }
    }

    private fun existingApp(): FirebaseApp? =
        runCatching { FirebaseApp.getInstance(APP_NAME) }.getOrNull()
}

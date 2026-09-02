package com.sentinel.sync

import android.content.Context
import java.util.UUID

/**
 * 기기 식별/소유자 식별의 로컬 저장.
 *
 * - [instanceId] : 이 설치의 안정적 하드 식별자(하드 바인딩용).
 * - [localOwnerId] : 태블릿(Master)이 자기 소유자 공간으로 쓰는 안정 ID.
 * - [pairedOwnerId] : 폰(Monitor)이 페어링으로 연결한 태블릿의 ownerId.
 *
 * (실기 정식판에선 Firebase Auth의 uid를 ownerId로 대체 예정. 지금은 로컬 UUID.)
 */
class IdentityStore(context: Context) {
    private val prefs = context.applicationContext.getSharedPreferences("sentinel_identity", Context.MODE_PRIVATE)

    val instanceId: String
        get() = prefs.getString(KEY_INSTANCE, null) ?: UUID.randomUUID().toString().also {
            prefs.edit().putString(KEY_INSTANCE, it).apply()
        }

    val localOwnerId: String
        get() = prefs.getString(KEY_OWNER, null) ?: UUID.randomUUID().toString().also {
            prefs.edit().putString(KEY_OWNER, it).apply()
        }

    var pairedOwnerId: String?
        get() = prefs.getString(KEY_PAIRED, null)
        set(value) { prefs.edit().putString(KEY_PAIRED, value).apply() }

    companion object {
        private const val KEY_INSTANCE = "instance_id"
        private const val KEY_OWNER = "owner_id"
        private const val KEY_PAIRED = "paired_owner_id"
    }
}

/** 페어링 코드 생성/도우미. */
object Pairing {
    private const val ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789" // 혼동 문자 제외
    private const val LEN = 6

    fun newCode(): String = buildString {
        repeat(LEN) { append(ALPHABET.random()) }
    }
}

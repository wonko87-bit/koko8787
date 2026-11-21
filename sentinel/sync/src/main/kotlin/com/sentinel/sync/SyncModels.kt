package com.sentinel.sync

/** 태블릿의 실시간 상태(폰이 구독). 항상 1문서로 덮어쓴다. */
data class LiveState(
    val status: String = "LOCKED",       // LOCKED | ACTIVE | EXPIRED
    val screenOn: Boolean = false,
    val currentApp: String? = null,      // 앱 라벨
    val currentPackage: String? = null,
    val sessionEndsAt: Long? = null,     // ACTIVE일 때 만료 epoch millis
    val updatedAt: Long = 0L,
) {
    fun toMap(): Map<String, Any?> = mapOf(
        "status" to status,
        "screenOn" to screenOn,
        "currentApp" to currentApp,
        "currentPackage" to currentPackage,
        "sessionEndsAt" to sessionEndsAt,
        "updatedAt" to updatedAt,
    )

    companion object {
        fun fromMap(m: Map<String, Any?>): LiveState = LiveState(
            status = m["status"] as? String ?: "LOCKED",
            screenOn = m["screenOn"] as? Boolean ?: false,
            currentApp = m["currentApp"] as? String,
            currentPackage = m["currentPackage"] as? String,
            sessionEndsAt = (m["sessionEndsAt"] as? Number)?.toLong(),
            updatedAt = (m["updatedAt"] as? Number)?.toLong() ?: 0L,
        )
    }
}

/** 세션 요약(폰 히스토리). */
data class SessionDto(
    val id: String = "",
    val grantedMillis: Long = 0L,
    val startedAt: Long = 0L,
    val approvedBy: String = "",
    val endedAt: Long? = null,
    val endReason: String? = null,
) {
    fun toMap(): Map<String, Any?> = mapOf(
        "id" to id,
        "grantedMillis" to grantedMillis,
        "startedAt" to startedAt,
        "approvedBy" to approvedBy,
        "endedAt" to endedAt,
        "endReason" to endReason,
    )

    companion object {
        fun fromMap(m: Map<String, Any?>): SessionDto = SessionDto(
            id = m["id"] as? String ?: "",
            grantedMillis = (m["grantedMillis"] as? Number)?.toLong() ?: 0L,
            startedAt = (m["startedAt"] as? Number)?.toLong() ?: 0L,
            approvedBy = m["approvedBy"] as? String ?: "",
            endedAt = (m["endedAt"] as? Number)?.toLong(),
            endReason = m["endReason"] as? String,
        )
    }
}

/** 이벤트(폰 로그 추출용). */
data class EventDto(
    val ts: Long = 0L,
    val type: String = "",
    val pkg: String? = null,
    val label: String? = null,
) {
    fun toMap(): Map<String, Any?> = mapOf(
        "ts" to ts, "type" to type, "pkg" to pkg, "label" to label,
    )

    companion object {
        fun fromMap(m: Map<String, Any?>): EventDto = EventDto(
            ts = (m["ts"] as? Number)?.toLong() ?: 0L,
            type = m["type"] as? String ?: "",
            pkg = m["pkg"] as? String,
            label = m["label"] as? String,
        )
    }
}

/** 바인딩된 기기(모니터/마스터). */
data class DeviceDto(
    val instanceId: String = "",
    val role: String = "",   // master | monitor
    val name: String = "",
    val boundAt: Long = 0L,
)

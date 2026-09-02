package com.sentinel.tablet.logging.db

import androidx.room.Entity
import androidx.room.Index
import androidx.room.PrimaryKey

/**
 * 한 세션의 요약 레코드. [SessionEngine]의 세션과 1:1.
 * 시작 시 insert, 종료 시 [endedAt]/[endReason] 업데이트.
 */
@Entity(tableName = "sessions")
data class SessionRecord(
    @PrimaryKey val id: String,
    val grantedMillis: Long,
    val startedAt: Long,
    val approvedBy: String,        // FINGERPRINT | PASSWORD
    val endedAt: Long? = null,
    val endReason: String? = null, // EXPIRED | MANUAL
)

/**
 * append-only 이벤트. 세션에 속하며 시각순으로 쌓인다.
 * [type]은 [EventType] 이름, [pkg]/[label]은 앱 전면 이벤트에서만 채워진다.
 */
@Entity(
    tableName = "events",
    indices = [Index("sessionId"), Index("ts")],
)
data class EventRecord(
    @PrimaryKey(autoGenerate = true) val id: Long = 0,
    val sessionId: String,
    val ts: Long,
    val type: String,
    val pkg: String? = null,
    val label: String? = null,
)

enum class EventType {
    SESSION_START,
    SESSION_END,
    APP_FOREGROUND,
    SCREEN_ON,
    SCREEN_OFF,
    UNLOCK,
}

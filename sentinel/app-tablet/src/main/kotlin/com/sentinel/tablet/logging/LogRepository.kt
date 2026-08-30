package com.sentinel.tablet.logging

import android.content.Context
import com.sentinel.core.Session
import com.sentinel.tablet.logging.db.EventRecord
import com.sentinel.tablet.logging.db.EventType
import com.sentinel.tablet.logging.db.LogDao
import com.sentinel.tablet.logging.db.LogDatabase
import com.sentinel.tablet.logging.db.SessionRecord
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.launch

/**
 * 로그 저장 파사드. 쓰기는 fire-and-forget(내부 IO 스코프)로 순서를 보장하기 위해
 * 단일 [CoroutineScope]에서 직렬 실행하지 않고, 이벤트 자체가 ts를 담으므로 순서는 ts로 복원한다.
 *
 * 읽기(뷰어/추출)는 [Flow] 또는 suspend 스냅샷으로 제공.
 */
class LogRepository private constructor(private val dao: LogDao) {

    private val scope = CoroutineScope(SupervisorJob() + Dispatchers.IO)

    // ---- 쓰기 ----

    fun onSessionStarted(session: Session) = scope.launch {
        dao.upsertSession(
            SessionRecord(
                id = session.id,
                grantedMillis = session.grantedMillis,
                startedAt = session.startedAt,
                approvedBy = session.approvedBy.name,
            ),
        )
        dao.insertEvent(EventRecord(sessionId = session.id, ts = session.startedAt, type = EventType.SESSION_START.name))
    }

    fun onSessionEnded(session: Session) = scope.launch {
        val endedAt = session.endedAt ?: System.currentTimeMillis()
        val reason = session.endReason?.name ?: "EXPIRED"
        dao.insertEvent(EventRecord(sessionId = session.id, ts = endedAt, type = EventType.SESSION_END.name))
        dao.closeSession(session.id, endedAt, reason)
    }

    fun logAppForeground(sessionId: String, ts: Long, pkg: String, label: String?) = scope.launch {
        dao.insertEvent(
            EventRecord(sessionId = sessionId, ts = ts, type = EventType.APP_FOREGROUND.name, pkg = pkg, label = label),
        )
    }

    fun logEvent(sessionId: String, ts: Long, type: EventType) = scope.launch {
        dao.insertEvent(EventRecord(sessionId = sessionId, ts = ts, type = type.name))
    }

    // ---- 읽기 ----

    fun observeSessions(): Flow<List<SessionRecord>> = dao.observeSessions()
    fun observeEvents(sessionId: String): Flow<List<EventRecord>> = dao.observeEvents(sessionId)

    suspend fun allSessions(): List<SessionRecord> = dao.allSessions()
    suspend fun allEvents(): List<EventRecord> = dao.allEvents()
    suspend fun eventsFor(sessionId: String): List<EventRecord> = dao.eventsFor(sessionId)

    companion object {
        @Volatile
        private var instance: LogRepository? = null

        fun get(context: Context): LogRepository =
            instance ?: synchronized(this) {
                instance ?: LogRepository(LogDatabase.get(context).logDao()).also { instance = it }
            }
    }
}

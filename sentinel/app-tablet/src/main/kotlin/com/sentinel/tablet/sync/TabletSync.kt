package com.sentinel.tablet.sync

import android.content.Context
import com.sentinel.core.Session
import com.sentinel.sync.EventDto
import com.sentinel.sync.IdentityStore
import com.sentinel.sync.LiveState
import com.sentinel.sync.Pairing
import com.sentinel.sync.SessionDto
import com.sentinel.sync.SyncClient
import com.sentinel.sync.SyncProvider
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.launch

/**
 * 태블릿(Master) → 클라우드 단방향 동기화. Firebase 미설정이면 [enabled]=false로 전부 no-op.
 * 로컬 기록(Room)이 원본이고, 여기선 폰 모니터링용 미러만 밀어 올린다.
 */
class TabletSync private constructor(context: Context) {

    private val sync: SyncClient = SyncProvider.get(context)
    private val identity = IdentityStore(context)
    private val scope = CoroutineScope(SupervisorJob() + Dispatchers.IO)

    val enabled: Boolean get() = sync.enabled
    val ownerId: String get() = identity.localOwnerId

    // 현재 실시간 상태(부분 업데이트 후 통째로 push).
    @Volatile
    private var live = LiveState()

    private fun push(update: LiveState.() -> LiveState) {
        if (!enabled) return
        val next = live.update().copy(updatedAt = System.currentTimeMillis())
        live = next
        scope.launch { runCatching { sync.pushLiveState(ownerId, next) } }
    }

    fun sessionStarted(session: Session) {
        if (!enabled) return
        scope.launch { runCatching { sync.upsertSession(ownerId, session.toDto()) } }
        push { copy(status = "ACTIVE", sessionEndsAt = session.endsAt) }
    }

    fun sessionEnded(session: Session) {
        if (!enabled) return
        scope.launch { runCatching { sync.upsertSession(ownerId, session.toDto()) } }
        push { copy(status = "LOCKED", sessionEndsAt = null, currentApp = null, currentPackage = null) }
    }

    fun appForeground(sessionId: String, ts: Long, pkg: String, label: String?) {
        if (!enabled) return
        scope.launch {
            runCatching {
                sync.addEvent(ownerId, sessionId, EventDto(ts = ts, type = "APP_FOREGROUND", pkg = pkg, label = label))
            }
        }
        push { copy(currentApp = label ?: pkg, currentPackage = pkg) }
    }

    fun screenEvent(sessionId: String, ts: Long, type: String) {
        if (!enabled) return
        scope.launch { runCatching { sync.addEvent(ownerId, sessionId, EventDto(ts = ts, type = type)) } }
        when (type) {
            "SCREEN_ON", "UNLOCK" -> push { copy(screenOn = true) }
            "SCREEN_OFF" -> push { copy(screenOn = false) }
        }
    }

    /** 페어링 코드 생성 후 등록. 실패 시 null. */
    suspend fun generatePairingCode(): String? {
        if (!enabled) return null
        val code = Pairing.newCode()
        return runCatching {
            sync.registerPairingCode(ownerId, code)
            code
        }.getOrNull()
    }

    private fun Session.toDto() = SessionDto(
        id = id,
        grantedMillis = grantedMillis,
        startedAt = startedAt,
        approvedBy = approvedBy.name,
        endedAt = endedAt,
        endReason = endReason?.name,
    )

    companion object {
        @Volatile
        private var instance: TabletSync? = null

        fun get(context: Context): TabletSync =
            instance ?: synchronized(this) {
                instance ?: TabletSync(context.applicationContext).also { instance = it }
            }
    }
}

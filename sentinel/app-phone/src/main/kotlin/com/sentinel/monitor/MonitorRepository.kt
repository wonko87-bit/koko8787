package com.sentinel.monitor

import android.content.Context
import android.os.Build
import com.sentinel.sync.EventDto
import com.sentinel.sync.IdentityStore
import com.sentinel.sync.LiveState
import com.sentinel.sync.SessionDto
import com.sentinel.sync.SyncClient
import com.sentinel.sync.SyncProvider
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.first
import kotlinx.coroutines.flow.flowOf

/**
 * 폰(Monitor) 읽기 전용 데이터 소스. Firebase 미설정이면 [enabled]=false.
 */
class MonitorRepository private constructor(context: Context) {

    private val sync: SyncClient = SyncProvider.get(context)
    private val identity = IdentityStore(context)

    val enabled: Boolean get() = sync.enabled
    val pairedOwnerId: String? get() = identity.pairedOwnerId
    val isPaired: Boolean get() = pairedOwnerId != null

    fun observeLive(): Flow<LiveState?> =
        pairedOwnerId?.let { sync.observeLiveState(it) } ?: flowOf(null)

    fun observeSessions(): Flow<List<SessionDto>> =
        pairedOwnerId?.let { sync.observeSessions(it) } ?: flowOf(emptyList())

    suspend fun events(sessionId: String): List<EventDto> =
        pairedOwnerId?.let { sync.eventsFor(it, sessionId) } ?: emptyList()

    /** 추출용 1회 스냅샷(현재 구독 값). */
    suspend fun sessionsSnapshot(): List<SessionDto> = observeSessions().first()

    /** 페어링 코드로 태블릿과 연결. 성공 시 true. */
    suspend fun pair(code: String): Boolean {
        if (!enabled) return false
        val ownerId = runCatching { sync.resolvePairingCode(code.trim().uppercase()) }.getOrNull() ?: return false
        identity.pairedOwnerId = ownerId
        runCatching { sync.bindDevice(ownerId, identity.instanceId, "monitor", deviceName()) }
        return true
    }

    fun unpair() {
        identity.pairedOwnerId = null
    }

    private fun deviceName(): String = "${Build.MANUFACTURER} ${Build.MODEL}"

    companion object {
        @Volatile
        private var instance: MonitorRepository? = null

        fun get(context: Context): MonitorRepository =
            instance ?: synchronized(this) {
                instance ?: MonitorRepository(context.applicationContext).also { instance = it }
            }
    }
}

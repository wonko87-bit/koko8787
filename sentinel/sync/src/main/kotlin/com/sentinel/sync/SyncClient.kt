package com.sentinel.sync

import android.content.Context
import com.google.firebase.firestore.FirebaseFirestore
import com.google.firebase.firestore.Query
import kotlinx.coroutines.channels.awaitClose
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.callbackFlow
import kotlinx.coroutines.flow.flowOf
import kotlinx.coroutines.tasks.await

/**
 * 태블릿(쓰기)·폰(읽기) 공용 동기화 클라이언트. Firebase 설정이 없으면 [Disabled]가 주입되어
 * 모든 호출이 안전하게 no-op/빈 값이 된다.
 */
interface SyncClient {
    val enabled: Boolean

    // ---- 태블릿(Master) 쓰기 ----
    suspend fun pushLiveState(ownerId: String, state: LiveState)
    suspend fun upsertSession(ownerId: String, session: SessionDto)
    suspend fun addEvent(ownerId: String, sessionId: String, event: EventDto)
    suspend fun registerPairingCode(ownerId: String, code: String)
    suspend fun bindDevice(ownerId: String, instanceId: String, role: String, name: String)

    // ---- 폰(Monitor) 읽기 ----
    fun observeLiveState(ownerId: String): Flow<LiveState?>
    fun observeSessions(ownerId: String): Flow<List<SessionDto>>
    suspend fun eventsFor(ownerId: String, sessionId: String): List<EventDto>
    suspend fun resolvePairingCode(code: String): String?
}

object SyncProvider {
    @Volatile
    private var instance: SyncClient? = null

    fun get(context: Context): SyncClient =
        instance ?: synchronized(this) {
            instance ?: build(context).also { instance = it }
        }

    private fun build(context: Context): SyncClient {
        val fs = FirebaseGate.firestore(context) ?: return DisabledSyncClient
        return FirestoreSyncClient(fs)
    }
}

/** Firebase 미설정 시 폴백. */
object DisabledSyncClient : SyncClient {
    override val enabled = false
    override suspend fun pushLiveState(ownerId: String, state: LiveState) {}
    override suspend fun upsertSession(ownerId: String, session: SessionDto) {}
    override suspend fun addEvent(ownerId: String, sessionId: String, event: EventDto) {}
    override suspend fun registerPairingCode(ownerId: String, code: String) {}
    override suspend fun bindDevice(ownerId: String, instanceId: String, role: String, name: String) {}
    override fun observeLiveState(ownerId: String): Flow<LiveState?> = flowOf(null)
    override fun observeSessions(ownerId: String): Flow<List<SessionDto>> = flowOf(emptyList())
    override suspend fun eventsFor(ownerId: String, sessionId: String): List<EventDto> = emptyList()
    override suspend fun resolvePairingCode(code: String): String? = null
}

private class FirestoreSyncClient(private val db: FirebaseFirestore) : SyncClient {

    override val enabled = true

    private fun liveDoc(owner: String) = db.document("owners/$owner/live/state")
    private fun sessions(owner: String) = db.collection("owners/$owner/sessions")
    private fun events(owner: String, sid: String) = db.collection("owners/$owner/sessions/$sid/events")

    override suspend fun pushLiveState(ownerId: String, state: LiveState) {
        liveDoc(ownerId).set(state.toMap()).await()
    }

    override suspend fun upsertSession(ownerId: String, session: SessionDto) {
        sessions(ownerId).document(session.id).set(session.toMap()).await()
    }

    override suspend fun addEvent(ownerId: String, sessionId: String, event: EventDto) {
        events(ownerId, sessionId).add(event.toMap()).await()
    }

    override suspend fun registerPairingCode(ownerId: String, code: String) {
        db.collection("pairing").document(code)
            .set(mapOf("ownerId" to ownerId, "createdAt" to System.currentTimeMillis())).await()
    }

    override suspend fun bindDevice(ownerId: String, instanceId: String, role: String, name: String) {
        db.document("owners/$ownerId/devices/$instanceId")
            .set(DeviceDto(instanceId, role, name, System.currentTimeMillis()).let {
                mapOf("instanceId" to it.instanceId, "role" to it.role, "name" to it.name, "boundAt" to it.boundAt)
            }).await()
    }

    override fun observeLiveState(ownerId: String): Flow<LiveState?> = callbackFlow {
        val reg = liveDoc(ownerId).addSnapshotListener { snap, err ->
            if (err != null) { trySend(null); return@addSnapshotListener }
            trySend(snap?.data?.let(LiveState::fromMap))
        }
        awaitClose { reg.remove() }
    }

    override fun observeSessions(ownerId: String): Flow<List<SessionDto>> = callbackFlow {
        val reg = sessions(ownerId).orderBy("startedAt", Query.Direction.DESCENDING)
            .addSnapshotListener { snap, err ->
                if (err != null) { trySend(emptyList()); return@addSnapshotListener }
                val list = snap?.documents?.mapNotNull { d -> d.data?.let(SessionDto::fromMap) } ?: emptyList()
                trySend(list)
            }
        awaitClose { reg.remove() }
    }

    override suspend fun eventsFor(ownerId: String, sessionId: String): List<EventDto> {
        val snap = events(ownerId, sessionId).orderBy("ts", Query.Direction.ASCENDING).get().await()
        return snap.documents.mapNotNull { d -> d.data?.let(EventDto::fromMap) }
    }

    override suspend fun resolvePairingCode(code: String): String? {
        val snap = db.collection("pairing").document(code).get().await()
        return snap.getString("ownerId")
    }
}

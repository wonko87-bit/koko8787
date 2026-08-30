package com.sentinel.tablet.logging.db

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import kotlinx.coroutines.flow.Flow

@Dao
interface LogDao {

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun upsertSession(session: SessionRecord)

    @Query(
        "UPDATE sessions SET endedAt = :endedAt, endReason = :reason WHERE id = :id",
    )
    suspend fun closeSession(id: String, endedAt: Long, reason: String)

    @Insert
    suspend fun insertEvent(event: EventRecord)

    @Query("SELECT * FROM sessions ORDER BY startedAt DESC")
    fun observeSessions(): Flow<List<SessionRecord>>

    @Query("SELECT * FROM events WHERE sessionId = :sessionId ORDER BY ts ASC")
    fun observeEvents(sessionId: String): Flow<List<EventRecord>>

    @Query("SELECT * FROM sessions ORDER BY startedAt ASC")
    suspend fun allSessions(): List<SessionRecord>

    @Query("SELECT * FROM events ORDER BY ts ASC")
    suspend fun allEvents(): List<EventRecord>

    @Query("SELECT * FROM events WHERE sessionId = :sessionId ORDER BY ts ASC")
    suspend fun eventsFor(sessionId: String): List<EventRecord>
}

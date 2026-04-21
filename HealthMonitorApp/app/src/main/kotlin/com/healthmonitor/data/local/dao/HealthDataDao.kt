package com.healthmonitor.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import com.healthmonitor.data.local.entities.HealthDataEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface HealthDataDao {

    @Query("SELECT * FROM health_data WHERE dateStr = :dateStr")
    suspend fun getByDate(dateStr: String): HealthDataEntity?

    @Query("SELECT * FROM health_data WHERE dateStr >= :startDateStr ORDER BY dateStr ASC")
    fun observeFrom(startDateStr: String): Flow<List<HealthDataEntity>>

    @Query("SELECT * FROM health_data WHERE dateStr >= :startDateStr ORDER BY dateStr ASC")
    suspend fun getFrom(startDateStr: String): List<HealthDataEntity>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun upsert(entity: HealthDataEntity)

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun upsertAll(entities: List<HealthDataEntity>)
}

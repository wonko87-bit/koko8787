package com.healthmonitor.data.local.entities

import androidx.room.Entity
import androidx.room.PrimaryKey
import com.healthmonitor.domain.models.HealthData
import java.time.LocalDate

@Entity(tableName = "health_data")
data class HealthDataEntity(
    @PrimaryKey val dateStr: String,
    val steps: Long,
    val sleepMinutes: Long,
    val activeCaloriesKcal: Double,
    val hydrationLiters: Double,
    val lastSyncTimestamp: Long
)

fun HealthDataEntity.toDomain() = HealthData(
    date = LocalDate.parse(dateStr),
    steps = steps,
    sleepMinutes = sleepMinutes,
    activeCaloriesKcal = activeCaloriesKcal,
    hydrationLiters = hydrationLiters
)

fun HealthData.toEntity(syncTime: Long = System.currentTimeMillis()) = HealthDataEntity(
    dateStr = date.toString(),
    steps = steps,
    sleepMinutes = sleepMinutes,
    activeCaloriesKcal = activeCaloriesKcal,
    hydrationLiters = hydrationLiters,
    lastSyncTimestamp = syncTime
)

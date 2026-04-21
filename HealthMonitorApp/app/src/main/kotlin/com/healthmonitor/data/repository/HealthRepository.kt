package com.healthmonitor.data.repository

import com.healthmonitor.data.health.HealthConnectManager
import com.healthmonitor.data.local.dao.HealthDataDao
import com.healthmonitor.data.local.entities.toDomain
import com.healthmonitor.data.local.entities.toEntity
import com.healthmonitor.domain.models.HealthData
import java.time.LocalDate
import javax.inject.Inject
import javax.inject.Singleton

@Singleton
class HealthRepository @Inject constructor(
    private val healthConnectManager: HealthConnectManager,
    private val healthDataDao: HealthDataDao
) {
    suspend fun hasAllPermissions() = healthConnectManager.hasAllPermissions()

    suspend fun syncWeeklyData(): List<HealthData> {
        val today = LocalDate.now()
        val results = (6 downTo 0).map { daysAgo ->
            val date = today.minusDays(daysAgo.toLong())
            val fresh = healthConnectManager.readDataForDate(date)
            healthDataDao.upsert(fresh.toEntity())
            fresh
        }
        return results
    }

    suspend fun getWeeklyData(): List<HealthData> {
        val startDate = LocalDate.now().minusDays(6)
        val cached = healthDataDao.getFrom(startDate.toString())
        if (cached.size == 7) return cached.map { it.toDomain() }
        // Cache miss: fetch from Health Connect
        return syncWeeklyData()
    }
}

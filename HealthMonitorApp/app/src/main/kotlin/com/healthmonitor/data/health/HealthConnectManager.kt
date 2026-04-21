package com.healthmonitor.data.health

import android.content.Context
import androidx.health.connect.client.HealthConnectClient
import androidx.health.connect.client.permission.HealthPermission
import androidx.health.connect.client.records.ActiveCaloriesBurnedRecord
import androidx.health.connect.client.records.HydrationRecord
import androidx.health.connect.client.records.SleepSessionRecord
import androidx.health.connect.client.records.StepsRecord
import androidx.health.connect.client.request.AggregateRequest
import androidx.health.connect.client.request.ReadRecordsRequest
import androidx.health.connect.client.time.TimeRangeFilter
import com.healthmonitor.domain.models.HealthData
import dagger.hilt.android.qualifiers.ApplicationContext
import java.time.Duration
import java.time.LocalDate
import java.time.ZoneId
import javax.inject.Inject
import javax.inject.Singleton

@Singleton
class HealthConnectManager @Inject constructor(
    @ApplicationContext private val context: Context
) {
    private val client: HealthConnectClient by lazy {
        HealthConnectClient.getOrCreate(context)
    }

    companion object {
        val REQUIRED_PERMISSIONS = setOf(
            HealthPermission.getReadPermission(StepsRecord::class),
            HealthPermission.getReadPermission(SleepSessionRecord::class),
            HealthPermission.getReadPermission(ActiveCaloriesBurnedRecord::class),
            HealthPermission.getReadPermission(HydrationRecord::class)
        )
    }

    fun getSdkStatus(): Int = HealthConnectClient.getSdkStatus(context)

    suspend fun hasAllPermissions(): Boolean =
        client.permissionController.getGrantedPermissions()
            .containsAll(REQUIRED_PERMISSIONS)

    suspend fun readDataForDate(date: LocalDate): HealthData {
        val zoneId = ZoneId.systemDefault()
        val dayStart = date.atStartOfDay(zoneId).toInstant()
        val dayEnd = date.plusDays(1).atStartOfDay(zoneId).toInstant()
        // Sleep is captured from 18:00 previous day to 12:00 current day
        val sleepStart = date.minusDays(1).atTime(18, 0).atZone(zoneId).toInstant()
        val sleepEnd = date.atTime(12, 0).atZone(zoneId).toInstant()

        val steps = readSteps(dayStart, dayEnd)
        val sleepMinutes = readSleepMinutes(sleepStart, sleepEnd)
        val activeCalories = readActiveCalories(dayStart, dayEnd)
        val hydration = readHydration(dayStart, dayEnd)

        return HealthData(
            date = date,
            steps = steps,
            sleepMinutes = sleepMinutes,
            activeCaloriesKcal = activeCalories,
            hydrationLiters = hydration
        )
    }

    private suspend fun readSteps(startTime: java.time.Instant, endTime: java.time.Instant): Long {
        val response = client.aggregate(
            AggregateRequest(
                metrics = setOf(StepsRecord.COUNT_TOTAL),
                timeRangeFilter = TimeRangeFilter.between(startTime, endTime)
            )
        )
        return response[StepsRecord.COUNT_TOTAL] ?: 0L
    }

    private suspend fun readSleepMinutes(startTime: java.time.Instant, endTime: java.time.Instant): Long {
        val response = client.readRecords(
            ReadRecordsRequest(
                recordType = SleepSessionRecord::class,
                timeRangeFilter = TimeRangeFilter.between(startTime, endTime)
            )
        )
        return response.records.sumOf { record ->
            Duration.between(record.startTime, record.endTime).toMinutes()
        }
    }

    private suspend fun readActiveCalories(startTime: java.time.Instant, endTime: java.time.Instant): Double {
        val response = client.aggregate(
            AggregateRequest(
                metrics = setOf(ActiveCaloriesBurnedRecord.ACTIVE_CALORIES_TOTAL),
                timeRangeFilter = TimeRangeFilter.between(startTime, endTime)
            )
        )
        return response[ActiveCaloriesBurnedRecord.ACTIVE_CALORIES_TOTAL]?.inKilocalories ?: 0.0
    }

    private suspend fun readHydration(startTime: java.time.Instant, endTime: java.time.Instant): Double {
        val response = client.aggregate(
            AggregateRequest(
                metrics = setOf(HydrationRecord.VOLUME_TOTAL),
                timeRangeFilter = TimeRangeFilter.between(startTime, endTime)
            )
        )
        return response[HydrationRecord.VOLUME_TOTAL]?.inLiters ?: 0.0
    }
}

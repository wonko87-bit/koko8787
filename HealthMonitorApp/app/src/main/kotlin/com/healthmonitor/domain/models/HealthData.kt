package com.healthmonitor.domain.models

import java.time.LocalDate

data class HealthData(
    val date: LocalDate,
    val steps: Long,
    val sleepMinutes: Long,
    val activeCaloriesKcal: Double,
    val hydrationLiters: Double
) {
    val sleepHours: Int get() = (sleepMinutes / 60).toInt()
    val sleepRemainingMinutes: Int get() = (sleepMinutes % 60).toInt()
    val stepsProgressFraction: Float get() = (steps / STEP_GOAL.toFloat()).coerceIn(0f, 1f)
    val hydrationProgressFraction: Float get() = (hydrationLiters / HYDRATION_GOAL_LITERS).toFloat().coerceIn(0f, 1f)

    companion object {
        const val STEP_GOAL = 10_000L
        const val HYDRATION_GOAL_LITERS = 2.0
    }
}

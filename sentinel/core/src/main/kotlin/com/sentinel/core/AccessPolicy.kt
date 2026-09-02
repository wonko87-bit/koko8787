package com.sentinel.core

/**
 * D3: 자동 부여 규칙. 실사용 기본값은 **수동 부여**이므로 [enabled] = false.
 * 기능은 미리 구현해 두고, 나중에 켜기만 하면 스케줄/일일예산 기반 자동 활성화가 동작한다.
 *
 * 순수 로직만 담는다. "지금 자동으로 열어줄 수 있는가?"를 [PolicyEvaluator]가 판정.
 */
data class AccessPolicy(
    val enabled: Boolean = false,
    /** 요일별 허용 시간창. 비어 있으면 시간창 제약 없음. */
    val windows: List<TimeWindow> = emptyList(),
    /** 하루 총 사용 예산(ms). null이면 무제한. */
    val dailyBudgetMillis: Long? = null,
    /** 자동 부여 시 1회 세션 길이(ms). */
    val autoGrantMillis: Long = 30 * 60_000L,
) {
    companion object {
        val DISABLED = AccessPolicy()
    }
}

/**
 * 요일 + 하루 중 분 단위 구간. [days]는 [java.time] 없이 쓰기 위해 1(월)~7(일) ISO 정수.
 * [startMinute]/[endMinute]는 자정 기준 분(0~1440).
 */
data class TimeWindow(
    val days: Set<Int>,
    val startMinute: Int,
    val endMinute: Int,
) {
    fun contains(isoDayOfWeek: Int, minuteOfDay: Int): Boolean =
        isoDayOfWeek in days && minuteOfDay in startMinute until endMinute
}

/** 자동 부여 가능 여부 판정 결과. */
sealed interface AutoGrantDecision {
    /** [millis]만큼 자동 부여 가능. */
    data class Grant(val millis: Long) : AutoGrantDecision

    /** 자동 부여 불가 + 사람이 읽을 사유. */
    data class Deny(val reason: String) : AutoGrantDecision
}

object PolicyEvaluator {
    /**
     * @param policy 현재 정책
     * @param isoDayOfWeek 1(월)~7(일)
     * @param minuteOfDay 0~1439
     * @param usedTodayMillis 오늘 이미 사용한 시간(ms)
     */
    fun evaluate(
        policy: AccessPolicy,
        isoDayOfWeek: Int,
        minuteOfDay: Int,
        usedTodayMillis: Long,
    ): AutoGrantDecision {
        if (!policy.enabled) return AutoGrantDecision.Deny("자동 부여 비활성(수동 모드)")

        if (policy.windows.isNotEmpty()) {
            val inWindow = policy.windows.any { it.contains(isoDayOfWeek, minuteOfDay) }
            if (!inWindow) return AutoGrantDecision.Deny("허용 시간창이 아님")
        }

        val budget = policy.dailyBudgetMillis
        val grant = if (budget == null) {
            policy.autoGrantMillis
        } else {
            val left = (budget - usedTodayMillis).coerceAtLeast(0L)
            if (left <= 0L) return AutoGrantDecision.Deny("오늘 예산 소진")
            minOf(policy.autoGrantMillis, left)
        }
        return AutoGrantDecision.Grant(Session.clampGrant(grant))
    }
}

package com.sentinel.core

/**
 * 한 번의 "부여 → 사용 → 종료" 세션. 순수 데이터 + 시간 계산만 담고,
 * 안드로이드 의존이 전혀 없어 단위 테스트가 쉽다.
 *
 * 모든 시각은 epoch millis (UTC) 기준.
 */
data class Session(
    val id: String,
    val grantedMillis: Long,
    val startedAt: Long,
    val approvedBy: ApprovalMethod,
    val endedAt: Long? = null,
    val endReason: EndReason? = null,
) {
    /** 이 세션이 만료되어야 하는 시각. */
    val endsAt: Long get() = startedAt + grantedMillis

    val isFinished: Boolean get() = endedAt != null

    /** [now] 기준 남은 시간(ms). 음수면 0으로 클램프. */
    fun remainingMillis(now: Long): Long = (endsAt - now).coerceAtLeast(0L)

    /** [now] 기준 만료 여부. */
    fun isExpiredAt(now: Long): Boolean = now >= endsAt

    /** 0.0~1.0 진행률 (HUD 게이지용). grantedMillis가 0이면 1.0. */
    fun progress(now: Long): Float {
        if (grantedMillis <= 0L) return 1f
        val elapsed = (now - startedAt).coerceIn(0L, grantedMillis)
        return elapsed.toFloat() / grantedMillis.toFloat()
    }

    fun finished(now: Long, reason: EndReason): Session =
        copy(endedAt = now, endReason = reason)

    companion object {
        const val MIN_GRANT_MILLIS: Long = 60_000L          // 최소 1분
        const val MAX_GRANT_MILLIS: Long = 12 * 3_600_000L  // 최대 12시간(안전장치)

        fun clampGrant(millis: Long): Long =
            millis.coerceIn(MIN_GRANT_MILLIS, MAX_GRANT_MILLIS)
    }
}

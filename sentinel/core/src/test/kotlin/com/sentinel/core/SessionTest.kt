package com.sentinel.core

import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertFalse
import kotlin.test.assertTrue

class SessionTest {

    private fun session(start: Long = 0L, grantMin: Long = 30): Session =
        Session(
            id = "s1",
            grantedMillis = grantMin * 60_000L,
            startedAt = start,
            approvedBy = ApprovalMethod.FINGERPRINT,
        )

    @Test
    fun `remaining counts down and clamps at zero`() {
        val s = session(grantMin = 30)
        assertEquals(30 * 60_000L, s.remainingMillis(0L))
        assertEquals(10 * 60_000L, s.remainingMillis(20 * 60_000L))
        assertEquals(0L, s.remainingMillis(30 * 60_000L))
        assertEquals(0L, s.remainingMillis(999 * 60_000L)) // 음수 클램프
    }

    @Test
    fun `expiry boundary is inclusive`() {
        val s = session(grantMin = 30)
        assertFalse(s.isExpiredAt(30 * 60_000L - 1))
        assertTrue(s.isExpiredAt(30 * 60_000L))
    }

    @Test
    fun `progress goes 0 to 1`() {
        val s = session(grantMin = 10)
        assertEquals(0f, s.progress(0L))
        assertEquals(0.5f, s.progress(5 * 60_000L))
        assertEquals(1f, s.progress(10 * 60_000L))
        assertEquals(1f, s.progress(99 * 60_000L)) // 초과해도 1로 클램프
    }

    @Test
    fun `grant is clamped to min and max`() {
        assertEquals(Session.MIN_GRANT_MILLIS, Session.clampGrant(0L))
        assertEquals(Session.MIN_GRANT_MILLIS, Session.clampGrant(30_000L))
        assertEquals(Session.MAX_GRANT_MILLIS, Session.clampGrant(Long.MAX_VALUE))
    }

    @Test
    fun `finished records reason and time`() {
        val s = session().finished(now = 12345L, reason = EndReason.MANUAL)
        assertTrue(s.isFinished)
        assertEquals(12345L, s.endedAt)
        assertEquals(EndReason.MANUAL, s.endReason)
    }
}

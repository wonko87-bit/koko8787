package com.sentinel.core

import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class PolicyEvaluatorTest {

    @Test
    fun `disabled policy always denies`() {
        val d = PolicyEvaluator.evaluate(AccessPolicy.DISABLED, 1, 600, 0L)
        assertTrue(d is AutoGrantDecision.Deny)
    }

    @Test
    fun `outside window denies`() {
        val policy = AccessPolicy(
            enabled = true,
            windows = listOf(TimeWindow(days = setOf(1, 2, 3, 4, 5), startMinute = 19 * 60, endMinute = 20 * 60)),
        )
        // 화요일 18:30 → 창 밖
        val d = PolicyEvaluator.evaluate(policy, 2, 18 * 60 + 30, 0L)
        assertTrue(d is AutoGrantDecision.Deny)
    }

    @Test
    fun `inside window grants configured length`() {
        val policy = AccessPolicy(
            enabled = true,
            windows = listOf(TimeWindow(days = setOf(2), startMinute = 19 * 60, endMinute = 20 * 60)),
            autoGrantMillis = 30 * 60_000L,
        )
        val d = PolicyEvaluator.evaluate(policy, 2, 19 * 60 + 10, 0L)
        assertEquals(AutoGrantDecision.Grant(30 * 60_000L), d)
    }

    @Test
    fun `daily budget caps the grant and denies when spent`() {
        val policy = AccessPolicy(
            enabled = true,
            dailyBudgetMillis = 40 * 60_000L,
            autoGrantMillis = 30 * 60_000L,
        )
        // 이미 20분 사용 → 남은 20분만 부여
        val d1 = PolicyEvaluator.evaluate(policy, 3, 600, 20 * 60_000L)
        assertEquals(AutoGrantDecision.Grant(20 * 60_000L), d1)

        // 예산 소진
        val d2 = PolicyEvaluator.evaluate(policy, 3, 600, 40 * 60_000L)
        assertTrue(d2 is AutoGrantDecision.Deny)
    }
}

package com.sentinel.core

/**
 * 태블릿의 3가지 상태.
 *
 * ```
 *            지문 인증 + 시간 부여 + 승인
 *   LOCKED  ────────────────────────────▶  ACTIVE
 *     ▲                                        │
 *     │        시간 만료 / 수동 종료            │
 *     └────────────────────────────────────────┘
 * ```
 *
 * - [LOCKED]  : 기본값. 키오스크로 잠긴 벽돌 상태.
 * - [ACTIVE]  : 관리자가 부여한 시간 동안만 사용 가능. 모든 활동 기록.
 * - [EXPIRED] : 시간이 막 끝난 찰나의 전이 상태(세션 마감 처리 후 즉시 LOCKED로).
 */
enum class TabletState {
    LOCKED,
    ACTIVE,
    EXPIRED,
}

/** 활성화를 승인한 수단. */
enum class ApprovalMethod {
    FINGERPRINT,
    PASSWORD,
}

/** 세션이 끝난 이유. */
enum class EndReason {
    /** 부여 시간이 자연 만료됨. */
    EXPIRED,

    /** 관리자가 별도 PW로 조기 종료함. */
    MANUAL,
}

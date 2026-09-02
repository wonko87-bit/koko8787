package com.sentinel.tablet.session

import com.sentinel.core.ApprovalMethod
import com.sentinel.core.EndReason
import com.sentinel.core.Session
import com.sentinel.core.TabletState
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import java.util.UUID

/**
 * 앱 전역 상태의 단일 원천(single source of truth).
 *
 * 순수 상태 전이만 담당하고, 타이머 틱은 [LockService]가 [tick]으로 밀어넣는다.
 * Compose UI는 [state]/[session]/[remainingMillis]를 구독한다.
 *
 * 부수효과(Device Owner 잠금, 화면 고정)는 [listener]로 위임 → 테스트 용이.
 */
object SessionEngine {

    interface Listener {
        /** LOCKED 진입: 키오스크 잠금 필요. */
        fun onLock()

        /** ACTIVE 진입: 잠금 해제, 사용 허용. [session] 로깅 시작. */
        fun onActivate(session: Session)

        /** 세션 종료(만료/수동): 로그 마감. */
        fun onSessionEnded(session: Session)
    }

    private val _state = MutableStateFlow(TabletState.LOCKED)
    val state: StateFlow<TabletState> = _state.asStateFlow()

    private val _session = MutableStateFlow<Session?>(null)
    val session: StateFlow<Session?> = _session.asStateFlow()

    private val _remainingMillis = MutableStateFlow(0L)
    val remainingMillis: StateFlow<Long> = _remainingMillis.asStateFlow()

    @Volatile
    private var listener: Listener? = null

    fun setListener(l: Listener?) {
        listener = l
    }

    /** 앱 시작/부팅 시 강제 잠금. */
    @Synchronized
    fun lock() {
        val active = _session.value
        if (active != null && !active.isFinished) {
            endInternal(active, System.currentTimeMillis(), EndReason.MANUAL)
        }
        _session.value = null
        _remainingMillis.value = 0L
        _state.value = TabletState.LOCKED
        listener?.onLock()
    }

    /**
     * 지문/PW 승인 후 [minutes]분 부여하며 ACTIVE로 전환.
     * @return 생성된 세션
     */
    @Synchronized
    fun activate(minutes: Int, method: ApprovalMethod, now: Long = System.currentTimeMillis()): Session {
        val granted = Session.clampGrant(minutes * 60_000L)
        val s = Session(
            id = UUID.randomUUID().toString(),
            grantedMillis = granted,
            startedAt = now,
            approvedBy = method,
        )
        _session.value = s
        _remainingMillis.value = s.remainingMillis(now)
        _state.value = TabletState.ACTIVE
        listener?.onActivate(s)
        return s
    }

    /** 타이머 틱. 만료되면 자동으로 EXPIRED→LOCKED 전이. */
    @Synchronized
    fun tick(now: Long = System.currentTimeMillis()) {
        val s = _session.value ?: return
        if (_state.value != TabletState.ACTIVE) return
        _remainingMillis.value = s.remainingMillis(now)
        if (s.isExpiredAt(now)) {
            _state.value = TabletState.EXPIRED
            endInternal(s, now, EndReason.EXPIRED)
            _session.value = null
            _remainingMillis.value = 0L
            _state.value = TabletState.LOCKED
            listener?.onLock()
        }
    }

    /** 관리자 별도 PW로 조기 종료. */
    @Synchronized
    fun endManually(now: Long = System.currentTimeMillis()) {
        val s = _session.value
        if (s != null && !s.isFinished) {
            _state.value = TabletState.EXPIRED
            endInternal(s, now, EndReason.MANUAL)
        }
        _session.value = null
        _remainingMillis.value = 0L
        _state.value = TabletState.LOCKED
        listener?.onLock()
    }

    private fun endInternal(s: Session, now: Long, reason: EndReason) {
        listener?.onSessionEnded(s.finished(now, reason))
    }
}

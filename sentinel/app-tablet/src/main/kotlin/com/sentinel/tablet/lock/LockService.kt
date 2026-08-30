package com.sentinel.tablet.lock

import android.app.Notification
import android.app.NotificationChannel
import android.app.NotificationManager
import android.app.PendingIntent
import android.content.Context
import android.content.Intent
import android.os.Build
import android.os.IBinder
import androidx.lifecycle.LifecycleService
import androidx.lifecycle.lifecycleScope
import com.sentinel.core.EndReason
import com.sentinel.core.Session
import com.sentinel.core.TabletState
import com.sentinel.tablet.R
import com.sentinel.tablet.admin.DeviceOwnerController
import com.sentinel.tablet.session.SessionEngine
import com.sentinel.tablet.ui.MainActivity
import kotlinx.coroutines.delay
import kotlinx.coroutines.flow.collectLatest
import kotlinx.coroutines.launch

/**
 * 상태를 살아있게 유지하고 만료를 감시하는 포그라운드 서비스.
 *
 * - 1초 틱으로 [SessionEngine.tick] 호출 → 남은시간 갱신/자동 만료.
 * - 만료(LOCKED 전이) 시 [MainActivity]를 전면으로 띄워 키오스크 재잠금.
 * - Device Owner 부수효과([SessionEngine.Listener])를 여기서 실행.
 *
 * (M2에서 이 서비스에 UsageStats/화면 이벤트 수집이 붙는다.)
 */
class LockService : LifecycleService(), SessionEngine.Listener {

    private lateinit var owner: DeviceOwnerController

    override fun onCreate() {
        super.onCreate()
        owner = DeviceOwnerController(this)
        owner.configureLockTaskAllowlist()
        SessionEngine.setListener(this)

        startForeground(NOTIF_ID, buildNotification(TabletState.LOCKED, 0L))

        // 1초 틱
        lifecycleScope.launch {
            while (true) {
                SessionEngine.tick()
                delay(1_000L)
            }
        }
        // 상태/남은시간 변화 → 알림 갱신
        lifecycleScope.launch {
            SessionEngine.state.collectLatest { updateNotification() }
        }
        lifecycleScope.launch {
            SessionEngine.remainingMillis.collectLatest { updateNotification() }
        }
    }

    override fun onStartCommand(intent: Intent?, flags: Int, startId: Int): Int {
        super.onStartCommand(intent, flags, startId)
        return START_STICKY
    }

    override fun onBind(intent: Intent): IBinder? {
        super.onBind(intent)
        return null
    }

    override fun onDestroy() {
        SessionEngine.setListener(null)
        super.onDestroy()
    }

    // ---- SessionEngine.Listener (부수효과) ----

    override fun onLock() {
        owner.enterLockdown()
        bringLockScreenToFront()
    }

    override fun onActivate(session: Session) {
        owner.exitLockdown()
        // TODO(M2): 세션 로깅 시작
    }

    override fun onSessionEnded(session: Session) {
        // TODO(M2): 세션 요약 로그 확정 + Firestore 동기화
        val reason = session.endReason ?: EndReason.EXPIRED
        android.util.Log.i(TAG, "session ${session.id} ended: $reason")
    }

    // ---- 알림 ----

    private fun updateNotification() {
        val nm = getSystemService(NotificationManager::class.java)
        nm.notify(NOTIF_ID, buildNotification(SessionEngine.state.value, SessionEngine.remainingMillis.value))
    }

    private fun buildNotification(state: TabletState, remaining: Long): Notification {
        ensureChannel()
        val text = when (state) {
            TabletState.ACTIVE -> "사용 중 · 남은 시간 ${formatRemaining(remaining)}"
            else -> "잠김 · 관리자 인증 필요"
        }
        val open = PendingIntent.getActivity(
            this, 0,
            Intent(this, MainActivity::class.java).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK),
            PendingIntent.FLAG_IMMUTABLE,
        )
        return Notification.Builder(this, CHANNEL_ID)
            .setContentTitle(getString(R.string.app_name))
            .setContentText(text)
            .setSmallIcon(R.drawable.ic_launcher)
            .setOngoing(true)
            .setContentIntent(open)
            .build()
    }

    private fun ensureChannel() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
            val nm = getSystemService(NotificationManager::class.java)
            if (nm.getNotificationChannel(CHANNEL_ID) == null) {
                nm.createNotificationChannel(
                    NotificationChannel(CHANNEL_ID, "Sentinel", NotificationManager.IMPORTANCE_LOW),
                )
            }
        }
    }

    private fun bringLockScreenToFront() {
        val intent = Intent(this, MainActivity::class.java)
            .addFlags(Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_SINGLE_TOP)
        startActivity(intent)
    }

    private fun formatRemaining(ms: Long): String {
        val totalSec = ms / 1000
        val m = totalSec / 60
        val s = totalSec % 60
        return "%d:%02d".format(m, s)
    }

    companion object {
        private const val TAG = "LockService"
        private const val CHANNEL_ID = "sentinel_lock"
        private const val NOTIF_ID = 1001

        fun start(context: Context) {
            val i = Intent(context, LockService::class.java)
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
                context.startForegroundService(i)
            } else {
                context.startService(i)
            }
        }
    }
}

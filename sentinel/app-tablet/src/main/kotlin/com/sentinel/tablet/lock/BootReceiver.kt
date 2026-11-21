package com.sentinel.tablet.lock

import android.content.BroadcastReceiver
import android.content.Context
import android.content.Intent
import com.sentinel.tablet.session.SessionEngine

/**
 * 부팅 완료 시 무조건 잠금 상태로 복귀시키고 감시 서비스를 되살린다.
 * (재부팅으로 잠금을 우회하려는 시도를 차단)
 */
class BootReceiver : BroadcastReceiver() {
    override fun onReceive(context: Context, intent: Intent) {
        when (intent.action) {
            Intent.ACTION_BOOT_COMPLETED,
            Intent.ACTION_LOCKED_BOOT_COMPLETED -> {
                SessionEngine.lock()
                LockService.start(context.applicationContext)
            }
        }
    }
}

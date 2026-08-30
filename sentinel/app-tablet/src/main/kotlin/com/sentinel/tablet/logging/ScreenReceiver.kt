package com.sentinel.tablet.logging

import android.content.BroadcastReceiver
import android.content.Context
import android.content.Intent
import android.content.IntentFilter
import com.sentinel.tablet.logging.db.EventType

/**
 * 화면 on/off 및 잠금해제(USER_PRESENT) 이벤트를 잡는 동적 리시버.
 * LockService가 세션 동안만 등록/해제한다(정적 등록 불가한 브로드캐스트라 코드로 등록).
 */
class ScreenReceiver(
    private val onEvent: (EventType) -> Unit,
) : BroadcastReceiver() {

    override fun onReceive(context: Context, intent: Intent) {
        when (intent.action) {
            Intent.ACTION_SCREEN_ON -> onEvent(EventType.SCREEN_ON)
            Intent.ACTION_SCREEN_OFF -> onEvent(EventType.SCREEN_OFF)
            Intent.ACTION_USER_PRESENT -> onEvent(EventType.UNLOCK)
        }
    }

    companion object {
        fun filter(): IntentFilter = IntentFilter().apply {
            addAction(Intent.ACTION_SCREEN_ON)
            addAction(Intent.ACTION_SCREEN_OFF)
            addAction(Intent.ACTION_USER_PRESENT)
        }
    }
}

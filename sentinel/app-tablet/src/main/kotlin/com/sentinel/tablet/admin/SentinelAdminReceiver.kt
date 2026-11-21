package com.sentinel.tablet.admin

import android.app.admin.DeviceAdminReceiver
import android.content.ComponentName
import android.content.Context
import android.content.Intent
import android.util.Log

/**
 * Device Owner(DPC) 진입점.
 *
 * 최초 1회 프로비저닝(계정이 하나도 없는 상태에서):
 * ```
 * adb shell dpm set-device-owner com.sentinel/com.sentinel.tablet.admin.SentinelAdminReceiver
 * ```
 * 지정 후에는 앱 삭제·강제종료·데이터 초기화가 막힌다.
 */
class SentinelAdminReceiver : DeviceAdminReceiver() {

    override fun onEnabled(context: Context, intent: Intent) {
        Log.i(TAG, "Device admin enabled")
    }

    override fun onDisabled(context: Context, intent: Intent) {
        Log.i(TAG, "Device admin disabled")
    }

    companion object {
        private const val TAG = "SentinelAdmin"

        fun componentName(context: Context): ComponentName =
            ComponentName(context.applicationContext, SentinelAdminReceiver::class.java)
    }
}

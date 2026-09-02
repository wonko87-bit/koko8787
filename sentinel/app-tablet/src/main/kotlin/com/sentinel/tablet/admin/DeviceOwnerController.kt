package com.sentinel.tablet.admin

import android.app.Activity
import android.app.admin.DevicePolicyManager
import android.content.Context
import android.os.Build
import android.os.UserManager
import android.util.Log

/**
 * DevicePolicyManager 래퍼. "사용불가"의 실제 강제력을 담당한다.
 *
 * Device Owner가 아니면(개발 중 일반 설치 등) 모든 호출이 조용히 no-op이 되어
 * 앱은 여전히 상태머신/UI 흐름을 시험할 수 있다. ([isDeviceOwner]로 구분)
 */
class DeviceOwnerController(context: Context) {

    private val appContext = context.applicationContext
    private val dpm =
        appContext.getSystemService(Context.DEVICE_POLICY_SERVICE) as DevicePolicyManager
    private val admin = SentinelAdminReceiver.componentName(appContext)

    val isDeviceOwner: Boolean
        get() = dpm.isDeviceOwnerApp(appContext.packageName)

    /** Sentinel 자신만 lock task 화이트리스트에 등록 (LOCKED 키오스크용). */
    fun configureLockTaskAllowlist() {
        if (!isDeviceOwner) return
        runCatching {
            dpm.setLockTaskPackages(admin, arrayOf(appContext.packageName))
        }.onFailure { Log.w(TAG, "setLockTaskPackages failed", it) }
    }

    /**
     * LOCKED 진입: 키오스크 잠금.
     * - 상태바/알림 비활성 → 설정·빠른설정 우회 차단
     * - 삭제 방지, 안전모드 부팅 제한
     * 실제 화면 고정(startLockTask)은 Activity 쪽에서 호출한다.
     */
    fun enterLockdown() {
        if (!isDeviceOwner) return
        runCatching {
            dpm.setStatusBarDisabled(admin, true)
            dpm.setUninstallBlocked(admin, appContext.packageName, true)
            addRestriction(UserManager.DISALLOW_SAFE_BOOT)
            addRestriction(UserManager.DISALLOW_ADD_USER)
            addRestriction(UserManager.DISALLOW_FACTORY_RESET)
        }.onFailure { Log.w(TAG, "enterLockdown failed", it) }
    }

    /** ACTIVE 진입: 사용 가능하도록 상태바 등 복원. 삭제 방지는 유지. */
    fun exitLockdown() {
        if (!isDeviceOwner) return
        runCatching {
            dpm.setStatusBarDisabled(admin, false)
        }.onFailure { Log.w(TAG, "exitLockdown failed", it) }
    }

    /** 화면 고정 대상 Activity에서 사용할 lock task 시작/종료 헬퍼. */
    fun startLockTask(activity: Activity) {
        runCatching { activity.startLockTask() }
            .onFailure { Log.w(TAG, "startLockTask failed", it) }
    }

    fun stopLockTask(activity: Activity) {
        runCatching { activity.stopLockTask() }
            .onFailure { Log.w(TAG, "stopLockTask failed", it) }
    }

    private fun addRestriction(key: String) {
        runCatching { dpm.addUserRestriction(admin, key) }
    }

    /** 프로비저닝 안내에 쓸 정확한 adb 명령. */
    fun provisioningCommand(): String =
        "adb shell dpm set-device-owner ${appContext.packageName}/" +
            SentinelAdminReceiver::class.java.name

    companion object {
        private const val TAG = "DeviceOwnerCtl"

        /** 이 기기가 lock task를 지원하는지(대부분의 실기 true). */
        fun lockTaskSupported(): Boolean = Build.VERSION.SDK_INT >= Build.VERSION_CODES.M
    }
}

package com.sentinel.tablet.logging

import android.app.AppOpsManager
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Process
import android.provider.Settings

/** PACKAGE_USAGE_STATS(사용 정보 접근)는 특수 권한 → 설정에서 사용자가 켜야 한다. */
object UsageAccess {

    fun isGranted(context: Context): Boolean {
        val appOps = context.getSystemService(Context.APP_OPS_SERVICE) as AppOpsManager
        val mode = appOps.unsafeCheckOpNoThrow(
            AppOpsManager.OPSTR_GET_USAGE_STATS,
            Process.myUid(),
            context.packageName,
        )
        return mode == AppOpsManager.MODE_ALLOWED
    }

    /** 사용 정보 접근 설정 화면을 연다. */
    fun openSettings(context: Context) {
        val intent = Intent(Settings.ACTION_USAGE_ACCESS_SETTINGS)
            .addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
        runCatching { context.startActivity(intent) }
    }
}

/** 패키지명 → 앱 이름. 실패 시 패키지명 그대로. 캐시로 반복 조회 비용 절감. */
class AppLabelResolver(context: Context) {
    private val pm: PackageManager = context.applicationContext.packageManager
    private val cache = HashMap<String, String>()

    fun label(pkg: String): String = cache.getOrPut(pkg) {
        runCatching {
            val info = pm.getApplicationInfo(pkg, 0)
            pm.getApplicationLabel(info).toString()
        }.getOrDefault(pkg)
    }
}

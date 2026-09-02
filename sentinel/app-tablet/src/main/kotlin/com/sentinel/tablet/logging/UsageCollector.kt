package com.sentinel.tablet.logging

import android.app.usage.UsageEvents
import android.app.usage.UsageStatsManager
import android.content.Context

/**
 * 전면(foreground) 앱 전환을 폴링으로 수집한다.
 *
 * [poll]을 주기적으로(예: 1초 틱) 호출하면, 직전 폴 이후의 UsageEvents를 훑어
 * 마지막으로 전면에 온 패키지가 바뀌었을 때만 [Change]를 돌려준다(중복 억제).
 *
 * PACKAGE_USAGE_STATS 권한이 없으면 항상 null.
 */
class UsageCollector(context: Context) {

    private val usm =
        context.getSystemService(Context.USAGE_STATS_SERVICE) as UsageStatsManager

    private var lastQueryTime: Long = System.currentTimeMillis()
    private var currentPackage: String? = null

    data class Change(val ts: Long, val pkg: String)

    /** 폴 시작 시점을 리셋(세션 시작 시 호출). */
    fun reset(now: Long = System.currentTimeMillis()) {
        lastQueryTime = now
        currentPackage = null
    }

    /**
     * @return 전면 앱이 바뀌었으면 [Change], 아니면 null.
     */
    fun poll(now: Long = System.currentTimeMillis()): Change? {
        val begin = lastQueryTime
        lastQueryTime = now
        if (now <= begin) return null

        val events = usm.queryEvents(begin, now)
        val e = UsageEvents.Event()
        var latestPkg: String? = null
        var latestTs: Long = begin
        while (events.hasNextEvent()) {
            events.getNextEvent(e)
            if (e.eventType == UsageEvents.Event.MOVE_TO_FOREGROUND) {
                latestPkg = e.packageName
                latestTs = e.timeStamp
            }
        }

        val pkg = latestPkg ?: return null
        if (pkg == currentPackage) return null
        currentPackage = pkg
        return Change(ts = latestTs, pkg = pkg)
    }
}

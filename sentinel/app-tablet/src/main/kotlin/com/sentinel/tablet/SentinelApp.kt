package com.sentinel.tablet

import android.app.Application
import com.sentinel.tablet.lock.LockService
import com.sentinel.tablet.session.SessionEngine

class SentinelApp : Application() {
    override fun onCreate() {
        super.onCreate()
        // 앱 프로세스 시작 시 항상 잠금부터.
        SessionEngine.lock()
        LockService.start(this)
    }
}

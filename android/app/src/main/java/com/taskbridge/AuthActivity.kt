package com.taskbridge

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.browser.customtabs.CustomTabsIntent
import androidx.lifecycle.lifecycleScope
import com.taskbridge.api.OAuthManager
import com.taskbridge.storage.TokenStore
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.util.UUID

/**
 * OAuth 플로우를 처리한다.
 * - 시작: provider("google"|"todoist") Intent extra로 전달
 * - 완료: taskbridge://auth/{google|todoist}?code=...&state=... 딥링크로 복귀
 */
class AuthActivity : AppCompatActivity() {

    private var pendingState: String? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        val provider = intent.getStringExtra("provider")
        if (provider != null) {
            startOAuth(provider)
            return
        }

        // 딥링크 콜백 처리
        handleCallback(intent)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleCallback(intent)
    }

    private fun startOAuth(provider: String) {
        val state = UUID.randomUUID().toString()
        pendingState = state

        val url = when (provider) {
            "google"  -> OAuthManager.googleAuthUrl(state)
            "todoist" -> OAuthManager.todoistAuthUrl(state)
            else -> { finish(); return }
        }

        CustomTabsIntent.Builder().build().launchUrl(this, Uri.parse(url))
    }

    private fun handleCallback(intent: Intent) {
        val uri = intent.data ?: return

        val code  = uri.getQueryParameter("code")  ?: return
        val state = uri.getQueryParameter("state")
        val path  = uri.path ?: ""

        lifecycleScope.launch {
            runCatching {
                withContext(Dispatchers.IO) {
                    when {
                        "/google"  in path -> OAuthManager.exchangeGoogleCode(this@AuthActivity, code)
                        "/todoist" in path -> OAuthManager.exchangeTodoistCode(this@AuthActivity, code)
                    }
                }
            }.onSuccess {
                val provider = if ("/google" in path) "Google Calendar" else "Todoist"
                Toast.makeText(this@AuthActivity, "✅ $provider 연결 완료!", Toast.LENGTH_SHORT).show()
            }.onFailure {
                Toast.makeText(this@AuthActivity, "❌ 연결 실패: ${it.message}", Toast.LENGTH_LONG).show()
            }

            startActivity(Intent(this@AuthActivity, MainActivity::class.java).addFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP))
            finish()
        }
    }
}

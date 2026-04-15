package com.taskbridge

import android.content.Intent
import android.os.Bundle
import android.view.View
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import com.taskbridge.api.GoogleCalApi
import com.taskbridge.api.OAuthManager
import com.taskbridge.api.TodoistApi
import com.taskbridge.databinding.ActivityMainBinding
import com.taskbridge.parser.InputParser
import com.taskbridge.parser.Target
import com.taskbridge.storage.TokenStore
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

class MainActivity : AppCompatActivity() {

    private lateinit var b: ActivityMainBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        b = ActivityMainBinding.inflate(layoutInflater)
        setContentView(b.root)

        b.btnSave.setOnClickListener { handleSave() }
        b.btnConnectGoogle.setOnClickListener  { startActivity(Intent(this, AuthActivity::class.java).putExtra("provider", "google")) }
        b.btnConnectTodoist.setOnClickListener { startActivity(Intent(this, AuthActivity::class.java).putExtra("provider", "todoist")) }

        // 입력창에서 Ctrl+Enter 대신 전송 버튼 사용
        b.etInput.setOnEditorActionListener { _, _, _ -> false }
    }

    override fun onResume() {
        super.onResume()
        updateAuthStatus()
    }

    private fun updateAuthStatus() {
        val google  = TokenStore.isGoogleConnected(this)
        val todoist = TokenStore.isTodoistConnected(this)

        b.btnConnectGoogle.text  = if (google)  "✓ Google Calendar 연결됨" else "Google Calendar 연결"
        b.btnConnectTodoist.text = if (todoist) "✓ Todoist 연결됨"         else "Todoist 연결"
        b.btnConnectGoogle.isEnabled  = !google
        b.btnConnectTodoist.isEnabled = !todoist

        if (google && todoist) b.authPanel.visibility = View.GONE
    }

    private fun handleSave() {
        val raw = b.etInput.text.toString().trim()
        if (raw.isEmpty()) { toast("내용을 입력해주세요."); return }

        val parsed  = InputParser.parse(raw)
        val content = parsed.text
        if (content.isEmpty()) { toast("내용이 없습니다."); return }

        val needGoogle  = parsed.target != Target.TODOIST
        val needTodoist = parsed.target != Target.CALENDAR

        if (needGoogle  && !TokenStore.isGoogleConnected(this))  { toast("Google Calendar 연결이 필요합니다."); return }
        if (needTodoist && !TokenStore.isTodoistConnected(this))  { toast("Todoist 연결이 필요합니다."); return }

        setLoading(true)

        lifecycleScope.launch {
            val errors = mutableListOf<String>()

            withContext(Dispatchers.IO) {
                if (needTodoist) {
                    runCatching {
                        TodoistApi.createTask(TokenStore.todoistAccessToken(this@MainActivity)!!, content)
                    }.onFailure { errors.add("Todoist: ${it.message}") }
                }

                if (needGoogle) {
                    val token = OAuthManager.validGoogleToken(this@MainActivity)
                    if (token == null) {
                        errors.add("Google 토큰 갱신 실패")
                    } else {
                        runCatching {
                            GoogleCalApi.createEvent(token, content)
                        }.onFailure { errors.add("Google Calendar: ${it.message}") }
                    }
                }
            }

            setLoading(false)

            if (errors.isEmpty()) {
                toast("✅ 저장되었습니다!")
                b.etInput.text?.clear()
            } else {
                toast("⚠️ ${errors.joinToString(" / ")}")
            }
        }
    }

    private fun setLoading(on: Boolean) {
        b.btnSave.isEnabled   = !on
        b.progressBar.visibility = if (on) View.VISIBLE else View.GONE
    }

    private fun toast(msg: String) = Toast.makeText(this, msg, Toast.LENGTH_SHORT).show()
}

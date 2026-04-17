package com.taskbridge.widget

import android.os.Bundle
import android.view.WindowManager
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import com.taskbridge.HelpDialog
import com.taskbridge.api.GoogleCalApi
import com.taskbridge.api.OAuthManager
import com.taskbridge.api.TodoistApi
import com.taskbridge.databinding.ActivityWidgetInputBinding
import com.taskbridge.parser.InputParser
import com.taskbridge.parser.Target
import com.taskbridge.storage.TokenStore
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

/**
 * 위젯 탭 시 표시되는 최소 입력 다이얼로그.
 * 저장 후 자동 종료.
 */
class WidgetInputActivity : AppCompatActivity() {

    private lateinit var b: ActivityWidgetInputBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        b = ActivityWidgetInputBinding.inflate(layoutInflater)
        setContentView(b.root)

        // 다이얼로그처럼 배경 어둡게
        window.addFlags(WindowManager.LayoutParams.FLAG_DIM_BEHIND)
        window.setDimAmount(0.5f)

        b.btnSend.setOnClickListener { handleSave() }
        b.btnCancel.setOnClickListener { finish() }
        b.btnHelp.setOnClickListener { HelpDialog.show(this) }
    }

    private fun handleSave() {
        val raw = b.etInput.text.toString().trim()
        if (raw.isEmpty()) { finish(); return }

        val parsed  = InputParser.parse(raw)
        val content = parsed.text.ifBlank { raw }

        b.btnSend.isEnabled = false

        lifecycleScope.launch {
            val errors = mutableListOf<String>()

            withContext(Dispatchers.IO) {
                if (parsed.target != Target.CALENDAR && TokenStore.isTodoistConnected(this@WidgetInputActivity)) {
                    runCatching {
                        TodoistApi.createTask(TokenStore.todoistAccessToken(this@WidgetInputActivity)!!, content)
                    }.onFailure { errors.add("Todoist 오류") }
                }

                if (parsed.target != Target.TODOIST && TokenStore.isGoogleConnected(this@WidgetInputActivity)) {
                    val token = OAuthManager.validGoogleToken(this@WidgetInputActivity)
                    if (token != null) {
                        runCatching {
                            GoogleCalApi.createEvent(token, content)
                        }.onFailure { errors.add("Calendar 오류") }
                    }
                }
            }

            Toast.makeText(
                this@WidgetInputActivity,
                if (errors.isEmpty()) "✅ 저장됨" else "⚠️ ${errors.joinToString()}",
                Toast.LENGTH_SHORT
            ).show()
            finish()
        }
    }
}

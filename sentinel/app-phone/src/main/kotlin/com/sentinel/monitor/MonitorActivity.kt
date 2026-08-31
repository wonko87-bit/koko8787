package com.sentinel.monitor

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.material3.darkColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.launch

class MonitorActivity : ComponentActivity() {

    private lateinit var repo: MonitorRepository

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        repo = MonitorRepository.get(this)
        setContent { MonitorTheme { App() } }
    }

    @Composable
    private fun App() {
        var paired by remember { mutableStateOf(repo.isPaired) }

        Surface(Modifier.fillMaxSize(), color = MaterialTheme.colorScheme.background) {
            when {
                !repo.enabled -> NotConfiguredScreen()
                !paired -> PairingEntryScreen(
                    repo = repo,
                    onPaired = { paired = true },
                )
                else -> DashboardScreen(
                    repo = repo,
                    onUnpair = { repo.unpair(); paired = false },
                    onExport = { exportCsv() },
                )
            }
        }
    }

    private fun exportCsv() {
        lifecycleScope.launch {
            runCatching {
                val file = MonitorCsvExporter.export(this@MonitorActivity, repo)
                MonitorCsvExporter.share(this@MonitorActivity, file)
            }
        }
    }
}

@Composable
private fun MonitorTheme(content: @Composable () -> Unit) {
    MaterialTheme(
        colorScheme = darkColorScheme(
            primary = Color(0xFF3BC0CC),
            secondary = Color(0xFF5FCE8A),
            error = Color(0xFFE2794F),
            background = Color(0xFF0E1116),
            surface = Color(0xFF161B22),
            onBackground = Color(0xFFE7EBF0),
            onSurface = Color(0xFFE7EBF0),
            onSurfaceVariant = Color(0xFFAAB3C0),
        ),
        content = content,
    )
}

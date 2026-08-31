package com.sentinel.monitor

import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.collectAsState
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableLongStateOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.sentinel.sync.LiveState
import com.sentinel.sync.SessionDto
import kotlinx.coroutines.delay
import kotlinx.coroutines.launch
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

private val clock = SimpleDateFormat("MM-dd HH:mm", Locale.KOREA)
private fun ts(ms: Long) = clock.format(Date(ms))
private fun mmss(ms: Long): String {
    val s = (ms / 1000).coerceAtLeast(0)
    return "%d:%02d".format(s / 60, s % 60)
}

@Composable
fun NotConfiguredScreen() {
    Box(Modifier.fillMaxSize().padding(28.dp), contentAlignment = Alignment.Center) {
        Column(horizontalAlignment = Alignment.CenterHorizontally) {
            Text("Sentinel Monitor", fontSize = 24.sp, color = MaterialTheme.colorScheme.primary)
            Spacer(Modifier.height(12.dp))
            Text(
                "Firebase가 설정되지 않았습니다.\nSENTINEL_FB_* 값을 넣고 다시 빌드하면 연동됩니다.",
                color = MaterialTheme.colorScheme.onSurfaceVariant, fontSize = 14.sp,
            )
        }
    }
}

@Composable
fun PairingEntryScreen(repo: MonitorRepository, onPaired: () -> Unit) {
    var code by remember { mutableStateOf("") }
    var error by remember { mutableStateOf<String?>(null) }
    var busy by remember { mutableStateOf(false) }
    val scope = rememberCoroutineScope()

    Box(Modifier.fillMaxSize().padding(24.dp), contentAlignment = Alignment.Center) {
        Card(colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface)) {
            Column(Modifier.padding(24.dp)) {
                Text("태블릿과 연동", fontSize = 22.sp, color = MaterialTheme.colorScheme.onSurface)
                Spacer(Modifier.height(4.dp))
                Text(
                    "태블릿의 '폰 연동'에서 생성한 코드를 입력하세요.",
                    fontSize = 13.sp, color = MaterialTheme.colorScheme.onSurfaceVariant,
                )
                Spacer(Modifier.height(20.dp))
                OutlinedTextField(
                    value = code,
                    onValueChange = { code = it.uppercase(); error = null },
                    label = { Text("페어링 코드") },
                    singleLine = true,
                    isError = error != null,
                    modifier = Modifier.fillMaxWidth(),
                )
                if (error != null) {
                    Spacer(Modifier.height(6.dp))
                    Text(error!!, color = MaterialTheme.colorScheme.error, fontSize = 13.sp)
                }
                Spacer(Modifier.height(20.dp))
                Button(
                    onClick = {
                        busy = true
                        scope.launch {
                            val ok = repo.pair(code)
                            busy = false
                            if (ok) onPaired() else error = "코드를 확인할 수 없습니다"
                        }
                    },
                    enabled = code.length >= 4 && !busy,
                    modifier = Modifier.fillMaxWidth(),
                ) { Text(if (busy) "연동 중…" else "연동") }
            }
        }
    }
}

@Composable
fun DashboardScreen(repo: MonitorRepository, onUnpair: () -> Unit, onExport: () -> Unit) {
    val live by repo.observeLive().collectAsState(initial = null)
    val sessions by repo.observeSessions().collectAsState(initial = emptyList())

    // 1초 카운트다운 틱
    var now by remember { mutableLongStateOf(System.currentTimeMillis()) }
    LaunchedEffect(Unit) {
        while (true) { now = System.currentTimeMillis(); delay(1_000L) }
    }

    Column(Modifier.fillMaxSize().padding(16.dp)) {
        Row(verticalAlignment = Alignment.CenterVertically) {
            Text("Sentinel Monitor", fontSize = 20.sp, color = MaterialTheme.colorScheme.primary)
            Spacer(Modifier.weight(1f))
            TextButton(onClick = onUnpair) { Text("연동 해제") }
        }
        Spacer(Modifier.height(12.dp))

        LiveCard(live, now)
        Spacer(Modifier.height(16.dp))

        Row(verticalAlignment = Alignment.CenterVertically) {
            Text("세션 기록", fontSize = 15.sp, color = MaterialTheme.colorScheme.onBackground)
            Spacer(Modifier.weight(1f))
            OutlinedButton(onClick = onExport, enabled = sessions.isNotEmpty()) { Text("CSV 추출") }
        }
        Spacer(Modifier.height(8.dp))

        if (sessions.isEmpty()) {
            Text("아직 기록이 없습니다.", color = MaterialTheme.colorScheme.onSurfaceVariant, fontSize = 13.sp)
        } else {
            LazyColumn(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                items(sessions, key = { it.id }) { s -> SessionRow(s) }
            }
        }
    }
}

@Composable
private fun LiveCard(live: LiveState?, now: Long) {
    val status = live?.status ?: "—"
    val (label, color) = when (status) {
        "ACTIVE" -> "사용 중" to MaterialTheme.colorScheme.secondary
        "LOCKED" -> "잠김" to MaterialTheme.colorScheme.error
        else -> "대기" to MaterialTheme.colorScheme.onSurfaceVariant
    }
    Card(
        modifier = Modifier.fillMaxWidth(),
        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
    ) {
        Column(Modifier.padding(18.dp)) {
            Row(verticalAlignment = Alignment.CenterVertically) {
                Text(label, fontSize = 20.sp, color = color)
                Spacer(Modifier.weight(1f))
                if (live?.screenOn == true) {
                    Text("화면 켜짐", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSurfaceVariant)
                }
            }
            if (status == "ACTIVE") {
                val remain = live?.sessionEndsAt?.let { it - now } ?: 0L
                Spacer(Modifier.height(10.dp))
                Text(
                    mmss(remain),
                    fontSize = 44.sp, fontFamily = FontFamily.Monospace,
                    color = MaterialTheme.colorScheme.onSurface,
                )
                Text("남은 시간", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSurfaceVariant)
                val app = live?.currentApp
                if (!app.isNullOrBlank()) {
                    Spacer(Modifier.height(10.dp))
                    Text("현재 앱 · $app", fontSize = 14.sp, color = MaterialTheme.colorScheme.onSurface)
                }
            }
            live?.updatedAt?.takeIf { it > 0 }?.let {
                Spacer(Modifier.height(10.dp))
                Text("업데이트 ${ts(it)}", fontSize = 11.sp, color = MaterialTheme.colorScheme.onSurfaceVariant)
            }
        }
    }
}

@Composable
private fun SessionRow(s: SessionDto) {
    Card(
        modifier = Modifier.fillMaxWidth(),
        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
    ) {
        Column(Modifier.padding(14.dp)) {
            Text(ts(s.startedAt), fontFamily = FontFamily.Monospace, color = MaterialTheme.colorScheme.onSurface)
            val status = when {
                s.endedAt == null -> "진행 중"
                s.endReason == "MANUAL" -> "수동 종료"
                else -> "만료 종료"
            }
            Text(
                "부여 ${s.grantedMillis / 60_000}분 · ${s.approvedBy} · $status",
                fontSize = 12.sp, color = MaterialTheme.colorScheme.onSurfaceVariant,
            )
        }
    }
}

package com.sentinel.tablet.ui

import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.Divider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.DisposableEffect
import androidx.compose.runtime.collectAsState
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalLifecycleOwner
import androidx.lifecycle.Lifecycle
import androidx.lifecycle.LifecycleEventObserver
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.sentinel.tablet.logging.LogRepository
import com.sentinel.tablet.logging.UsageAccess
import com.sentinel.tablet.logging.db.EventRecord
import com.sentinel.tablet.logging.db.EventType
import com.sentinel.tablet.logging.db.SessionRecord
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

private val clock = SimpleDateFormat("MM-dd HH:mm:ss", Locale.KOREA)
private fun t(ms: Long) = clock.format(Date(ms))
private fun mmss(ms: Long): String {
    val s = ms / 1000
    return "%d:%02d".format(s / 60, s % 60)
}

@Composable
fun LogScreen(
    repo: LogRepository,
    onBack: () -> Unit,
    onExport: () -> Unit,
) {
    val context = LocalContext.current
    val sessions by repo.observeSessions().collectAsState(initial = emptyList())
    // 설정에서 권한을 켜고 돌아오면 배너가 사라지도록 resume마다 재확인.
    var usageGranted by remember { mutableStateOf(UsageAccess.isGranted(context)) }
    val lifecycleOwner = LocalLifecycleOwner.current
    DisposableEffect(lifecycleOwner) {
        val observer = LifecycleEventObserver { _, event ->
            if (event == Lifecycle.Event.ON_RESUME) usageGranted = UsageAccess.isGranted(context)
        }
        lifecycleOwner.lifecycle.addObserver(observer)
        onDispose { lifecycleOwner.lifecycle.removeObserver(observer) }
    }

    Column(Modifier.fillMaxSize().padding(16.dp)) {
        Row(verticalAlignment = Alignment.CenterVertically) {
            TextButton(onClick = onBack) { Text("← 뒤로") }
            Spacer(Modifier.weight(1f))
            Text("로그", fontSize = 20.sp, color = MaterialTheme.colorScheme.onBackground)
            Spacer(Modifier.weight(1f))
            OutlinedButton(onClick = onExport, enabled = sessions.isNotEmpty()) { Text("CSV 추출") }
        }
        Spacer(Modifier.height(8.dp))

        if (!usageGranted) {
            Card(
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.errorContainer),
                modifier = Modifier.fillMaxWidth(),
            ) {
                Column(Modifier.padding(14.dp)) {
                    Text(
                        "사용 정보 접근 권한이 꺼져 있어 앱 사용 기록이 남지 않습니다.",
                        color = MaterialTheme.colorScheme.onErrorContainer, fontSize = 13.sp,
                    )
                    Spacer(Modifier.height(6.dp))
                    TextButton(onClick = {
                        UsageAccess.openSettings(context)
                    }) { Text("설정 열기") }
                }
            }
            Spacer(Modifier.height(8.dp))
        }

        if (sessions.isEmpty()) {
            Column(
                Modifier.fillMaxSize(), verticalArrangement = Arrangement.Center,
                horizontalAlignment = Alignment.CenterHorizontally,
            ) {
                Text("아직 기록된 세션이 없습니다.", color = MaterialTheme.colorScheme.onSurfaceVariant)
            }
        } else {
            LazyColumn(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                items(sessions, key = { it.id }) { s -> SessionCard(repo, s) }
            }
        }
    }
}

@Composable
private fun SessionCard(repo: LogRepository, s: SessionRecord) {
    var expanded by remember { mutableStateOf(false) }

    Card(
        modifier = Modifier.fillMaxWidth().clickable { expanded = !expanded },
        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
    ) {
        Column(Modifier.padding(14.dp)) {
            Row(verticalAlignment = Alignment.CenterVertically) {
                Column(Modifier.weight(1f)) {
                    Text(t(s.startedAt), fontFamily = FontFamily.Monospace, color = MaterialTheme.colorScheme.onSurface)
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
                Text(if (expanded) "▲" else "▼", color = MaterialTheme.colorScheme.onSurfaceVariant)
            }

            if (expanded) {
                val events by remember(s.id) { repo.observeEvents(s.id) }.collectAsState(initial = emptyList())
                Spacer(Modifier.height(10.dp))
                Divider(color = MaterialTheme.colorScheme.outline)
                Spacer(Modifier.height(10.dp))

                AppTotals(events, sessionEnd = s.endedAt)
                Spacer(Modifier.height(10.dp))
                Text("타임라인", fontSize = 12.sp, color = MaterialTheme.colorScheme.primary)
                Spacer(Modifier.height(4.dp))
                events.forEach { e -> EventRow(e) }
            }
        }
    }
}

@Composable
private fun AppTotals(events: List<EventRecord>, sessionEnd: Long?) {
    val totals = remember(events, sessionEnd) { computeAppTotals(events, sessionEnd) }
    if (totals.isEmpty()) {
        Text("앱 사용 기록 없음", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSurfaceVariant)
        return
    }
    Text("앱별 사용시간", fontSize = 12.sp, color = MaterialTheme.colorScheme.primary)
    Spacer(Modifier.height(4.dp))
    totals.forEach { (label, ms) ->
        Row(Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.SpaceBetween) {
            Text(label, fontSize = 13.sp, color = MaterialTheme.colorScheme.onSurface)
            Text(mmss(ms), fontSize = 13.sp, fontFamily = FontFamily.Monospace, color = MaterialTheme.colorScheme.onSurface)
        }
    }
}

@Composable
private fun EventRow(e: EventRecord) {
    val desc = when (EventType.valueOf(e.type)) {
        EventType.SESSION_START -> "세션 시작"
        EventType.SESSION_END -> "세션 종료"
        EventType.SCREEN_ON -> "화면 켜짐"
        EventType.SCREEN_OFF -> "화면 꺼짐"
        EventType.UNLOCK -> "잠금 해제"
        EventType.APP_FOREGROUND -> "앱 · ${e.label ?: e.pkg}"
    }
    Row(Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(10.dp)) {
        Text(t(e.ts), fontSize = 11.sp, fontFamily = FontFamily.Monospace, color = MaterialTheme.colorScheme.onSurfaceVariant)
        Text(desc, fontSize = 12.sp, color = MaterialTheme.colorScheme.onSurface)
    }
}

/** APP_FOREGROUND 이벤트를 연속 구간으로 보고 패키지별 총 사용시간(ms)을 계산. 내림차순. */
private fun computeAppTotals(events: List<EventRecord>, sessionEnd: Long?): List<Pair<String, Long>> {
    val fg = events.filter { it.type == EventType.APP_FOREGROUND.name }
    if (fg.isEmpty()) return emptyList()
    val bound = sessionEnd ?: events.maxOf { it.ts }
    val totals = LinkedHashMap<String, Long>()
    for (i in fg.indices) {
        val cur = fg[i]
        val next = if (i + 1 < fg.size) fg[i + 1].ts else bound
        val dur = (next - cur.ts).coerceAtLeast(0L)
        val key = cur.label ?: cur.pkg ?: "?"
        totals[key] = (totals[key] ?: 0L) + dur
    }
    return totals.entries.sortedByDescending { it.value }.map { it.key to it.value }
}

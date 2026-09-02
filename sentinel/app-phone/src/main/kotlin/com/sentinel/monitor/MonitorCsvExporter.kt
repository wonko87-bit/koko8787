package com.sentinel.monitor

import android.content.Context
import android.content.Intent
import androidx.core.content.FileProvider
import com.sentinel.sync.SessionDto
import java.io.File
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

/** 폰에서 클라우드 로그를 CSV로 추출·공유. */
object MonitorCsvExporter {

    private val fmt = SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.KOREA)

    suspend fun export(context: Context, repo: MonitorRepository): File {
        val sessions = repo.sessionsSnapshot()

        val sb = StringBuilder()
        sb.append("session_id,session_started,granted_min,approved_by,end_reason,")
        sb.append("event_time,event_epoch_ms,event_type,package,app_label\n")

        for (s in sessions) {
            val events = runCatching { repo.events(s.id) }.getOrDefault(emptyList())
            if (events.isEmpty()) {
                sb.append(sessionCols(s)).append(",,,,,\n")
                continue
            }
            for (e in events) {
                sb.append(sessionCols(s)).append(',')
                sb.append(csv(time(e.ts))).append(',')
                sb.append(e.ts).append(',')
                sb.append(csv(e.type)).append(',')
                sb.append(csv(e.pkg ?: "")).append(',')
                sb.append(csv(e.label ?: "")).append('\n')
            }
        }

        val dir = File(context.cacheDir, "exports").apply { mkdirs() }
        val file = File(dir, "sentinel-monitor-${System.currentTimeMillis()}.csv")
        file.writeText(sb.toString())
        return file
    }

    fun share(context: Context, file: File) {
        val uri = FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", file)
        val intent = Intent(Intent.ACTION_SEND).apply {
            type = "text/csv"
            putExtra(Intent.EXTRA_STREAM, uri)
            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        }
        runCatching {
            context.startActivity(
                Intent.createChooser(intent, "로그 CSV 내보내기").addFlags(Intent.FLAG_ACTIVITY_NEW_TASK),
            )
        }
    }

    private fun sessionCols(s: SessionDto): String =
        listOf(
            csv(s.id), csv(time(s.startedAt)), (s.grantedMillis / 60_000).toString(),
            csv(s.approvedBy), csv(s.endReason ?: ""),
        ).joinToString(",")

    private fun time(ms: Long) = fmt.format(Date(ms))

    private fun csv(v: String): String =
        if (v.contains(',') || v.contains('"') || v.contains('\n') || v.contains('\r')) {
            "\"" + v.replace("\"", "\"\"") + "\""
        } else v
}

package com.sentinel.tablet.logging

import android.content.Context
import android.content.Intent
import androidx.core.content.FileProvider
import com.sentinel.tablet.logging.db.SessionRecord
import java.io.File
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

/**
 * 세션·이벤트를 단일 평면 CSV로 추출하고 공유 시트로 내보낸다.
 * 각 이벤트 행에 소속 세션 정보를 붙여, 그대로 스프레드시트에서 분석 가능.
 */
object CsvExporter {

    private val fmt = SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.KOREA)

    /** @return 생성된 CSV 파일. */
    suspend fun export(context: Context, repo: LogRepository): File {
        val sessions = repo.allSessions().associateBy { it.id }
        val events = repo.allEvents()

        val sb = StringBuilder()
        sb.append("session_id,session_started,granted_min,approved_by,end_reason,")
        sb.append("event_time,event_epoch_ms,event_type,package,app_label\n")

        for (e in events) {
            val s: SessionRecord? = sessions[e.sessionId]
            sb.append(csv(e.sessionId)).append(',')
            sb.append(csv(s?.startedAt?.let(::time) ?: "")).append(',')
            sb.append(csv(s?.let { (it.grantedMillis / 60_000).toString() } ?: "")).append(',')
            sb.append(csv(s?.approvedBy ?: "")).append(',')
            sb.append(csv(s?.endReason ?: "")).append(',')
            sb.append(csv(time(e.ts))).append(',')
            sb.append(e.ts).append(',')
            sb.append(csv(e.type)).append(',')
            sb.append(csv(e.pkg ?: "")).append(',')
            sb.append(csv(e.label ?: "")).append('\n')
        }

        val dir = File(context.cacheDir, "exports").apply { mkdirs() }
        val file = File(dir, "sentinel-log-${System.currentTimeMillis()}.csv")
        file.writeText(sb.toString())
        return file
    }

    /** 파일을 공유 시트로 내보낸다(메일/드라이브/폰으로 전송 등). */
    fun share(context: Context, file: File) {
        val uri = FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", file)
        val intent = Intent(Intent.ACTION_SEND).apply {
            type = "text/csv"
            putExtra(Intent.EXTRA_STREAM, uri)
            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        }
        val chooser = Intent.createChooser(intent, "로그 CSV 내보내기")
            .addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
        runCatching { context.startActivity(chooser) }
    }

    private fun time(ms: Long): String = fmt.format(Date(ms))

    /** RFC-ish CSV escaping. */
    private fun csv(v: String): String =
        if (v.contains(',') || v.contains('"') || v.contains('\n')) {
            "\"" + v.replace("\"", "\"\"") + "\""
        } else v
}

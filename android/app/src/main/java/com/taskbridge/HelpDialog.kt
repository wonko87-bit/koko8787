package com.taskbridge

import android.content.Context
import androidx.appcompat.app.AlertDialog

object HelpDialog {

    private val CONTENT = """
📌 기본 사용법
• 입력 후 저장 → Todoist + Google Calendar 동시 저장
• /c 내용 → Google Calendar 에만 저장
• /t 내용 → Todoist 에만 저장

📅 날짜 · 시간
• 오늘 / 내일 / 모레
• 3월 15일 / 15일
• 오전 9시 / 오후 3시 30분 / 14:30

⏱ 시간 범위
• 오후 2시 ~ 4시 회의
• 오전 10시 ~ 오후 1시 미팅
• (시작 시간만 적으면 1시간 이벤트로 자동 생성)

🔁 반복 일정
• 매일 오전 9시 스탠드업
• 매주 월요일 오후 2시 팀미팅
• 매격주 수요일 1:1 미팅
• 매달 1일 월간 리뷰

🔔 리마인더
• !r15 → 15분 전 알림
• !r30 → 30분 전 알림
• !r60 → 1시간 전 알림
• 알림은 Todoist · Google Calendar 양쪽 모두 적용

🏷 Todoist 메타데이터
• #프로젝트 → 프로젝트 지정 (없으면 자동 생성)
• @레이블 → 레이블 추가
• !p1 긴급 / !p2 높음 / !p3 보통 / !p4 낮음

💡 조합 예시
• 매주 금요일 오후 5시 주간보고 #업무 !p2 !r30
• /c 내일 오전 10시 ~ 12시 발표 준비
• /t 오늘 보고서 초안 @중요 !p1
• 오후 3시 ~ 5시 외부 미팅 #영업 @미팅 !r30
    """.trimIndent()

    fun show(ctx: Context) {
        AlertDialog.Builder(ctx)
            .setTitle("📖 사용 가이드")
            .setMessage(CONTENT)
            .setPositiveButton("확인", null)
            .show()
    }
}

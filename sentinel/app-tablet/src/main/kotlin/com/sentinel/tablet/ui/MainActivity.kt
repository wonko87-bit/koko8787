package com.sentinel.tablet.ui

import android.os.Bundle
import androidx.activity.compose.setContent
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.runtime.collectAsState
import androidx.fragment.app.FragmentActivity
import com.sentinel.core.ApprovalMethod
import com.sentinel.core.TabletState
import com.sentinel.tablet.admin.DeviceOwnerController
import com.sentinel.tablet.security.BiometricGate
import com.sentinel.tablet.security.PasswordVault
import com.sentinel.tablet.session.SessionEngine

/**
 * 잠금 키오스크이자 홈 런처. 화면은 [SessionEngine] 상태로 완전히 구동된다.
 *
 * 키오스크 강제:
 * - LOCKED/EXPIRED → startLockTask (화면 고정, 이탈 불가)
 * - ACTIVE → stopLockTask + moveTaskToBack (태블릿 자유 사용, HUD는 앱 재실행 시)
 */
class MainActivity : FragmentActivity() {

    private lateinit var owner: DeviceOwnerController
    private lateinit var vault: PasswordVault

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        owner = DeviceOwnerController(this)
        vault = PasswordVault(this)

        setContent { com.sentinel.tablet.ui.theme.SentinelTheme { Root() } }
    }

    @androidx.compose.runtime.Composable
    private fun Root() {
        val state by SessionEngine.state.collectAsState()
        val remaining by SessionEngine.remainingMillis.collectAsState()

        var configured by remember { mutableStateOf(vault.isConfigured) }
        var adminAuthed by remember { mutableStateOf(false) }   // 활성화 화면 노출
        var approveViaPwMinutes by remember { mutableStateOf<Int?>(null) }
        var adminViaPw by remember { mutableStateOf(false) }     // 지문 대신 PW 관리자 진입
        var endingSession by remember { mutableStateOf(false) }
        var pwError by remember { mutableStateOf<String?>(null) }

        // 상태 전이 → 키오스크 제어 + 임시 UI 플래그 리셋
        LaunchedEffect(state) {
            when (state) {
                TabletState.LOCKED, TabletState.EXPIRED -> {
                    adminAuthed = false
                    approveViaPwMinutes = null
                    adminViaPw = false
                    endingSession = false
                    pwError = null
                    if (DeviceOwnerController.lockTaskSupported()) owner.startLockTask(this@MainActivity)
                }
                TabletState.ACTIVE -> {
                    owner.stopLockTask(this@MainActivity)
                    moveTaskToBack(true)
                }
            }
        }

        when {
            !configured -> PasswordSetupScreen(onSet = {
                vault.setPassword(it)
                configured = true
            })

            state == TabletState.ACTIVE -> {
                if (endingSession) {
                    PasswordPromptScreen(
                        title = "세션 종료",
                        error = pwError,
                        onSubmit = { pw ->
                            if (vault.verify(pw)) {
                                SessionEngine.endManually()
                            } else pwError = "PW가 일치하지 않습니다"
                        },
                        onCancel = { endingSession = false; pwError = null },
                    )
                } else {
                    ActiveHudScreen(
                        remainingText = formatClock(remaining),
                        onEndSession = { endingSession = true },
                    )
                }
            }

            // LOCKED / EXPIRED
            adminViaPw -> PasswordPromptScreen(
                title = "관리자 인증",
                error = pwError,
                onSubmit = { pw ->
                    if (vault.verify(pw)) {
                        adminViaPw = false; pwError = null; adminAuthed = true
                    } else pwError = "PW가 일치하지 않습니다"
                },
                onCancel = { adminViaPw = false; pwError = null },
            )

            adminAuthed && approveViaPwMinutes != null -> PasswordPromptScreen(
                title = "활성화 승인",
                error = pwError,
                onSubmit = { pw ->
                    if (vault.verify(pw)) {
                        SessionEngine.activate(approveViaPwMinutes!!, ApprovalMethod.PASSWORD)
                    } else pwError = "PW가 일치하지 않습니다"
                },
                onCancel = { approveViaPwMinutes = null; pwError = null },
            )

            adminAuthed -> ActivationScreen(
                onApproveFingerprint = { minutes ->
                    authFingerprint(
                        title = "활성화 승인",
                        subtitle = "$minutes 분 사용을 시작합니다",
                        onOk = { SessionEngine.activate(minutes, ApprovalMethod.FINGERPRINT) },
                    )
                },
                onApprovePassword = { minutes -> approveViaPwMinutes = minutes; pwError = null },
                onCancel = { adminAuthed = false },
            )

            else -> LockScreen(
                onAdmin = {
                    authFingerprint(
                        title = "관리자 인증",
                        subtitle = "지문으로 관리자 콘솔에 진입",
                        onOk = { adminAuthed = true },
                    )
                },
                onAdminPassword = { adminViaPw = true; pwError = null },
                biometricNote = biometricNote(),
            )
        }
    }

    private fun biometricNote(): String? =
        when (BiometricGate.availability(this)) {
            BiometricGate.Availability.AVAILABLE -> null
            BiometricGate.Availability.NONE_ENROLLED -> "등록된 지문이 없습니다 · PW로 진입하세요"
            BiometricGate.Availability.NO_HARDWARE -> "지문 하드웨어를 사용할 수 없습니다"
            BiometricGate.Availability.UNAVAILABLE -> "지문을 사용할 수 없습니다"
        }

    private fun authFingerprint(title: String, subtitle: String, onOk: () -> Unit) {
        if (BiometricGate.availability(this) != BiometricGate.Availability.AVAILABLE) return
        BiometricGate.prompt(
            activity = this,
            title = title,
            subtitle = subtitle,
            onSuccess = onOk,
            onFail = { _, _ -> /* 취소/오류: 아무 것도 안 함 (사용자 재시도) */ },
        )
    }

    private fun formatClock(ms: Long): String {
        val totalSec = ms / 1000
        return "%d:%02d".format(totalSec / 60, totalSec % 60)
    }
}

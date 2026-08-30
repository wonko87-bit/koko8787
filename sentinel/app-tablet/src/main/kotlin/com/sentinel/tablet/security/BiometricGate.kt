package com.sentinel.tablet.security

import androidx.biometric.BiometricManager
import androidx.biometric.BiometricManager.Authenticators.BIOMETRIC_STRONG
import androidx.biometric.BiometricPrompt
import androidx.core.content.ContextCompat
import androidx.fragment.app.FragmentActivity

/**
 * 관리자 지문 인증.
 *
 * 핵심: [BIOMETRIC_STRONG]만 허용한다. 이렇게 하면
 * - 패턴/PIN/비밀번호(device credential)로의 폴백이 발생하지 않는다 → "패턴 안 됨" 보장.
 * - Class 3(STRONG)만 받으므로, 얼굴(대개 Class 2)이 등록돼 있어도 배제되고 사실상 지문만 남는다.
 *   (기기 펌웨어의 생체 등급 분류에 의존 → 실기 1회 검증 필요)
 */
object BiometricGate {

    enum class Availability { AVAILABLE, NONE_ENROLLED, NO_HARDWARE, UNAVAILABLE }

    fun availability(activity: FragmentActivity): Availability {
        val bm = BiometricManager.from(activity)
        return when (bm.canAuthenticate(BIOMETRIC_STRONG)) {
            BiometricManager.BIOMETRIC_SUCCESS -> Availability.AVAILABLE
            BiometricManager.BIOMETRIC_ERROR_NONE_ENROLLED -> Availability.NONE_ENROLLED
            BiometricManager.BIOMETRIC_ERROR_NO_HARDWARE,
            BiometricManager.BIOMETRIC_ERROR_HW_UNAVAILABLE -> Availability.NO_HARDWARE
            else -> Availability.UNAVAILABLE
        }
    }

    /**
     * 지문 프롬프트를 띄운다.
     * @param onSuccess 인증 성공
     * @param onFail    사용자 취소 또는 하드웨어 오류(errorCode, 메시지)
     */
    fun prompt(
        activity: FragmentActivity,
        title: String,
        subtitle: String,
        onSuccess: () -> Unit,
        onFail: (code: Int, message: CharSequence) -> Unit,
    ) {
        val executor = ContextCompat.getMainExecutor(activity)
        val prompt = BiometricPrompt(
            activity,
            executor,
            object : BiometricPrompt.AuthenticationCallback() {
                override fun onAuthenticationSucceeded(result: BiometricPrompt.AuthenticationResult) {
                    onSuccess()
                }

                override fun onAuthenticationError(errorCode: Int, errString: CharSequence) {
                    onFail(errorCode, errString)
                }
                // onAuthenticationFailed(단순 불일치)는 프롬프트가 자체 재시도를 유도하므로 무시.
            },
        )

        val info = BiometricPrompt.PromptInfo.Builder()
            .setTitle(title)
            .setSubtitle(subtitle)
            .setAllowedAuthenticators(BIOMETRIC_STRONG) // 패턴/PIN 폴백 배제
            .setNegativeButtonText("취소")               // STRONG 전용이라 취소 버튼 필수
            .setConfirmationRequired(false)
            .build()

        prompt.authenticate(info)
    }
}

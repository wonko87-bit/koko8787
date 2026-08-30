package com.sentinel.tablet.ui

import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.layout.widthIn
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Fingerprint
import androidx.compose.material.icons.filled.Lock
import androidx.compose.material3.Button
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.FilterChip
import androidx.compose.material3.Icon
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.text.input.PasswordVisualTransformation
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp

/** 첫 실행 — 별도 PW 설정. */
@Composable
fun PasswordSetupScreen(onSet: (CharArray) -> Unit) {
    var pw by remember { mutableStateOf("") }
    var pw2 by remember { mutableStateOf("") }
    val ok = pw.length >= 4 && pw == pw2

    CenterCard(title = "별도 PW 설정", subtitle = "안드로이드 잠금과 다른, Sentinel 전용 비밀번호") {
        OutlinedTextField(
            value = pw, onValueChange = { pw = it }, label = { Text("PW (4자 이상)") },
            singleLine = true, visualTransformation = PasswordVisualTransformation(),
            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.NumberPassword),
            modifier = Modifier.fillMaxWidth(),
        )
        Spacer(Modifier.height(12.dp))
        OutlinedTextField(
            value = pw2, onValueChange = { pw2 = it }, label = { Text("PW 확인") },
            singleLine = true, visualTransformation = PasswordVisualTransformation(),
            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.NumberPassword),
            isError = pw2.isNotEmpty() && pw != pw2,
            modifier = Modifier.fillMaxWidth(),
        )
        Spacer(Modifier.height(20.dp))
        Button(onClick = { onSet(pw.toCharArray()) }, enabled = ok, modifier = Modifier.fillMaxWidth()) {
            Text("설정하고 시작")
        }
    }
}

/** LOCKED — 잠금 화면 + 관리자 진입 버튼. */
@Composable
fun LockScreen(onAdmin: () -> Unit, onAdminPassword: () -> Unit, biometricNote: String?) {
    Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
        Column(horizontalAlignment = Alignment.CenterHorizontally) {
            Icon(
                Icons.Filled.Lock, contentDescription = null,
                tint = MaterialTheme.colorScheme.error, modifier = Modifier.size(72.dp),
            )
            Spacer(Modifier.height(20.dp))
            Text("잠김", fontSize = 34.sp, color = MaterialTheme.colorScheme.onBackground)
            Spacer(Modifier.height(8.dp))
            Text(
                "관리자 인증 후에만 사용할 수 있습니다.",
                color = MaterialTheme.colorScheme.onSurfaceVariant,
            )
            Spacer(Modifier.height(40.dp))
            Button(
                onClick = onAdmin,
                colors = ButtonDefaults.buttonColors(containerColor = MaterialTheme.colorScheme.primary),
            ) {
                Icon(Icons.Filled.Fingerprint, contentDescription = null)
                Spacer(Modifier.width(10.dp))
                Text("관리자 지문 인증")
            }
            if (biometricNote != null) {
                Spacer(Modifier.height(16.dp))
                Text(biometricNote, color = MaterialTheme.colorScheme.error, fontSize = 13.sp)
            }
            Spacer(Modifier.height(12.dp))
            TextButton(onClick = onAdminPassword) { Text("지문 대신 별도 PW로 인증") }
        }
    }
}

/** 활성화 — 시간 부여 + 승인. */
@Composable
fun ActivationScreen(
    onApproveFingerprint: (Int) -> Unit,
    onApprovePassword: (Int) -> Unit,
    onCancel: () -> Unit,
) {
    var minutes by remember { mutableStateOf(30) }
    val presets = listOf(10, 20, 30, 60, 90)

    CenterCard(title = "사용 시간 부여", subtitle = "이 시간이 끝나면 자동으로 다시 잠깁니다") {
        Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
            presets.forEach { p ->
                FilterChip(
                    selected = minutes == p,
                    onClick = { minutes = p },
                    label = { Text("${p}분") },
                )
            }
        }
        Spacer(Modifier.height(16.dp))
        Row(verticalAlignment = Alignment.CenterVertically) {
            OutlinedButton(onClick = { minutes = (minutes - 5).coerceAtLeast(1) }) { Text("-5") }
            Spacer(Modifier.width(16.dp))
            Text(
                "$minutes 분",
                fontSize = 28.sp, fontFamily = FontFamily.Monospace,
                color = MaterialTheme.colorScheme.onSurface,
            )
            Spacer(Modifier.width(16.dp))
            OutlinedButton(onClick = { minutes += 5 }) { Text("+5") }
        }
        Spacer(Modifier.height(24.dp))
        Button(onClick = { onApproveFingerprint(minutes) }, modifier = Modifier.fillMaxWidth()) {
            Icon(Icons.Filled.Fingerprint, contentDescription = null)
            Spacer(Modifier.width(10.dp))
            Text("지문으로 승인 · 사용 시작")
        }
        Spacer(Modifier.height(8.dp))
        OutlinedButton(onClick = { onApprovePassword(minutes) }, modifier = Modifier.fillMaxWidth()) {
            Text("별도 PW로 승인")
        }
        Spacer(Modifier.height(4.dp))
        TextButton(onClick = onCancel, modifier = Modifier.fillMaxWidth()) { Text("취소") }
    }
}

/**
 * ACTIVE HUD — 평소엔 태블릿을 백그라운드로 보내 사용하게 하고,
 * 사용자가 앱을 다시 열었을 때만 이 화면이 보인다(남은시간 + 관리자 종료).
 */
@Composable
fun ActiveHudScreen(remainingText: String, onEndSession: () -> Unit) {
    Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
        Column(horizontalAlignment = Alignment.CenterHorizontally) {
            Text("사용 중", color = MaterialTheme.colorScheme.secondary, fontSize = 18.sp)
            Spacer(Modifier.height(12.dp))
            Text(
                remainingText,
                fontSize = 56.sp, fontFamily = FontFamily.Monospace,
                color = MaterialTheme.colorScheme.onBackground,
            )
            Spacer(Modifier.height(8.dp))
            Text("남은 시간", color = MaterialTheme.colorScheme.onSurfaceVariant)
            Spacer(Modifier.height(36.dp))
            OutlinedButton(onClick = onEndSession) { Text("관리자 종료 (별도 PW)") }
        }
    }
}

/** 별도 PW 입력 다이얼로그 대체용 인라인 카드. */
@Composable
fun PasswordPromptScreen(
    title: String,
    onSubmit: (CharArray) -> Unit,
    onCancel: () -> Unit,
    error: String? = null,
) {
    var pw by remember { mutableStateOf("") }
    CenterCard(title = title, subtitle = "Sentinel 별도 PW") {
        OutlinedTextField(
            value = pw, onValueChange = { pw = it }, label = { Text("PW") },
            singleLine = true, visualTransformation = PasswordVisualTransformation(),
            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.NumberPassword),
            isError = error != null, modifier = Modifier.fillMaxWidth(),
        )
        if (error != null) {
            Spacer(Modifier.height(6.dp))
            Text(error, color = MaterialTheme.colorScheme.error, fontSize = 13.sp)
        }
        Spacer(Modifier.height(20.dp))
        Button(onClick = { onSubmit(pw.toCharArray()) }, enabled = pw.isNotEmpty(), modifier = Modifier.fillMaxWidth()) {
            Text("확인")
        }
        Spacer(Modifier.height(4.dp))
        TextButton(onClick = onCancel, modifier = Modifier.fillMaxWidth()) { Text("취소") }
    }
}

@Composable
fun LoadingScreen() {
    Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
        CircularProgressIndicator()
    }
}

@Composable
private fun CenterCard(
    title: String,
    subtitle: String,
    content: @Composable () -> Unit,
) {
    Box(Modifier.fillMaxSize().padding(24.dp), contentAlignment = Alignment.Center) {
        Card(
            modifier = Modifier.widthIn(max = 420.dp).fillMaxWidth(),
            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
        ) {
            Column(Modifier.padding(24.dp)) {
                Text(title, fontSize = 22.sp, color = MaterialTheme.colorScheme.onSurface)
                Spacer(Modifier.height(4.dp))
                Text(subtitle, color = MaterialTheme.colorScheme.onSurfaceVariant, fontSize = 13.sp)
                Spacer(Modifier.height(20.dp))
                content()
            }
        }
    }
}

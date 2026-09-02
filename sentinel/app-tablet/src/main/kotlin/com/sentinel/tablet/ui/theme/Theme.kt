package com.sentinel.tablet.ui.theme

import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.ui.graphics.Color

// Sentinel 팔레트 — 슬레이트 지면 + 티일 강조, 잠금은 러스트.
private val Teal = Color(0xFF3BC0CC)
private val TealDark = Color(0xFF0E7C86)
private val Rust = Color(0xFFE2794F)
private val Go = Color(0xFF5FCE8A)

private val DarkColors = darkColorScheme(
    primary = Teal,
    onPrimary = Color(0xFF04252A),
    secondary = Go,
    error = Rust,
    background = Color(0xFF0E1116),
    onBackground = Color(0xFFE7EBF0),
    surface = Color(0xFF161B22),
    onSurface = Color(0xFFE7EBF0),
    surfaceVariant = Color(0xFF1C222B),
    onSurfaceVariant = Color(0xFFAAB3C0),
    outline = Color(0xFF2A313C),
)

private val LightColors = lightColorScheme(
    primary = TealDark,
    onPrimary = Color.White,
    secondary = Color(0xFF2E7D4F),
    error = Color(0xFFB5451F),
    background = Color(0xFFECEEF1),
    onBackground = Color(0xFF1A1E26),
    surface = Color(0xFFF7F8FA),
    onSurface = Color(0xFF1A1E26),
    outline = Color(0xFFD6DAE1),
)

@Composable
fun SentinelTheme(
    dark: Boolean = isSystemInDarkTheme(),
    content: @Composable () -> Unit,
) {
    MaterialTheme(
        colorScheme = if (dark) DarkColors else LightColors,
        content = content,
    )
}

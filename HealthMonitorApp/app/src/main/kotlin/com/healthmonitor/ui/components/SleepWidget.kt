package com.healthmonitor.ui.components

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Bedtime
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.healthmonitor.domain.models.HealthData
import com.healthmonitor.ui.theme.CardSurface
import com.healthmonitor.ui.theme.SleepColor
import com.healthmonitor.ui.theme.SleepDim
import com.healthmonitor.ui.theme.TextSecondary

private const val SLEEP_GOAL_HOURS = 8

@Composable
fun SleepWidget(data: HealthData, modifier: Modifier = Modifier) {
    val progressFraction = (data.sleepMinutes / (SLEEP_GOAL_HOURS * 60f)).coerceIn(0f, 1f)
    val quality = when {
        data.sleepMinutes >= 7 * 60 -> "양호"
        data.sleepMinutes >= 5 * 60 -> "보통"
        data.sleepMinutes > 0 -> "부족"
        else -> "데이터 없음"
    }

    Card(
        modifier = modifier,
        shape = RoundedCornerShape(16.dp),
        colors = CardDefaults.cardColors(containerColor = CardSurface)
    ) {
        Box(
            modifier = Modifier
                .fillMaxSize()
                .background(
                    Brush.verticalGradient(
                        colors = listOf(SleepDim.copy(alpha = 0.3f), CardSurface)
                    )
                )
                .padding(16.dp)
        ) {
            Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.spacedBy(6.dp)
                ) {
                    Icon(
                        Icons.Default.Bedtime,
                        contentDescription = null,
                        tint = SleepColor,
                        modifier = Modifier.size(18.dp)
                    )
                    Text(
                        text = "수면",
                        style = MaterialTheme.typography.titleMedium,
                        color = TextSecondary
                    )
                }

                if (data.sleepMinutes > 0) {
                    Row(
                        verticalAlignment = Alignment.Baseline,
                        horizontalArrangement = Arrangement.spacedBy(2.dp)
                    ) {
                        Text(
                            text = "${data.sleepHours}",
                            fontSize = 28.sp,
                            fontWeight = FontWeight.Bold,
                            color = SleepColor
                        )
                        Text(
                            text = "시간",
                            style = MaterialTheme.typography.bodyMedium,
                            color = TextSecondary
                        )
                        Spacer(Modifier.width(4.dp))
                        Text(
                            text = "${data.sleepRemainingMinutes}",
                            fontSize = 28.sp,
                            fontWeight = FontWeight.Bold,
                            color = SleepColor
                        )
                        Text(
                            text = "분",
                            style = MaterialTheme.typography.bodyMedium,
                            color = TextSecondary
                        )
                    }
                } else {
                    Text(
                        text = "—",
                        fontSize = 28.sp,
                        fontWeight = FontWeight.Bold,
                        color = SleepColor
                    )
                }

                Text(
                    text = "목표 ${SLEEP_GOAL_HOURS}시간 · $quality",
                    style = MaterialTheme.typography.labelSmall,
                    color = TextSecondary
                )

                LinearProgressIndicator(
                    progress = { progressFraction },
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(6.dp)
                        .clip(RoundedCornerShape(3.dp)),
                    color = SleepColor,
                    trackColor = SleepDim
                )
            }
        }
    }
}

package com.healthmonitor.ui.components

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.LocalFireDepartment
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
import com.healthmonitor.ui.theme.CaloriesColor
import com.healthmonitor.ui.theme.CaloriesDim
import com.healthmonitor.ui.theme.CardSurface
import com.healthmonitor.ui.theme.TextSecondary
import kotlin.math.roundToInt

private const val CALORIE_GOAL = 500.0

@Composable
fun CaloriesWidget(data: HealthData, modifier: Modifier = Modifier) {
    val progressFraction = (data.activeCaloriesKcal / CALORIE_GOAL).toFloat().coerceIn(0f, 1f)

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
                        colors = listOf(CaloriesDim.copy(alpha = 0.3f), CardSurface)
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
                        Icons.Default.LocalFireDepartment,
                        contentDescription = null,
                        tint = CaloriesColor,
                        modifier = Modifier.size(18.dp)
                    )
                    Text(
                        text = "활동 칼로리",
                        style = MaterialTheme.typography.titleMedium,
                        color = TextSecondary
                    )
                }

                Row(
                    verticalAlignment = Alignment.Baseline,
                    horizontalArrangement = Arrangement.spacedBy(2.dp)
                ) {
                    Text(
                        text = data.activeCaloriesKcal.roundToInt().toString(),
                        fontSize = 28.sp,
                        fontWeight = FontWeight.Bold,
                        color = CaloriesColor
                    )
                    Text(
                        text = "kcal",
                        style = MaterialTheme.typography.bodyMedium,
                        color = TextSecondary
                    )
                }

                Text(
                    text = "목표 ${CALORIE_GOAL.toInt()} kcal",
                    style = MaterialTheme.typography.labelSmall,
                    color = TextSecondary
                )

                LinearProgressIndicator(
                    progress = { progressFraction },
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(6.dp)
                        .clip(RoundedCornerShape(3.dp)),
                    color = CaloriesColor,
                    trackColor = CaloriesDim
                )

                Text(
                    text = "${(progressFraction * 100).toInt()}%",
                    style = MaterialTheme.typography.labelSmall,
                    color = CaloriesColor
                )
            }
        }
    }
}

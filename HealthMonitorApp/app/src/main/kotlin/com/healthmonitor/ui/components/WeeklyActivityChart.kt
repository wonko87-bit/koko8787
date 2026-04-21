package com.healthmonitor.ui.components

import androidx.compose.foundation.Canvas
import androidx.compose.foundation.layout.*
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.geometry.CornerRadius
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.geometry.Size
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.PathEffect
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import com.healthmonitor.domain.models.HealthData
import com.healthmonitor.ui.theme.Primary
import com.healthmonitor.ui.theme.TextSecondary
import java.time.LocalDate
import java.time.format.TextStyle
import java.util.Locale

@Composable
fun WeeklyActivityChart(
    weeklyData: List<HealthData>,
    modifier: Modifier = Modifier
) {
    if (weeklyData.isEmpty()) return

    val today = LocalDate.now()
    val maxSteps = maxOf(weeklyData.maxOf { it.steps }, HealthData.STEP_GOAL)
    val barColor = Primary
    val barDimColor = Color(0xFF37474F)
    val goalLineColor = Primary.copy(alpha = 0.35f)

    Column(modifier = modifier) {
        Canvas(
            modifier = Modifier
                .fillMaxWidth()
                .height(140.dp)
        ) {
            val count = weeklyData.size
            val horizontalPadding = 12.dp.toPx()
            val bottomPadding = 4.dp.toPx()
            val availableWidth = size.width - horizontalPadding * 2
            val slotWidth = availableWidth / count
            val barWidth = slotWidth * 0.55f
            val maxBarHeight = size.height - bottomPadding

            val goalFraction = HealthData.STEP_GOAL.toFloat() / maxSteps
            val goalY = size.height - goalFraction * maxBarHeight - bottomPadding

            // Dashed goal line
            drawLine(
                color = goalLineColor,
                start = Offset(horizontalPadding, goalY),
                end = Offset(size.width - horizontalPadding, goalY),
                strokeWidth = 1.5.dp.toPx(),
                pathEffect = PathEffect.dashPathEffect(floatArrayOf(10f, 6f))
            )

            weeklyData.forEachIndexed { index, data ->
                val fraction = data.steps.toFloat() / maxSteps
                val barHeight = (fraction * maxBarHeight).coerceAtLeast(4.dp.toPx())
                val x = horizontalPadding + index * slotWidth + (slotWidth - barWidth) / 2
                val y = size.height - barHeight - bottomPadding

                drawRoundRect(
                    color = if (data.date == today) barColor else barDimColor,
                    topLeft = Offset(x, y),
                    size = Size(barWidth, barHeight),
                    cornerRadius = CornerRadius(4.dp.toPx())
                )
            }
        }

        Spacer(Modifier.height(4.dp))

        // Day labels
        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.SpaceEvenly
        ) {
            weeklyData.forEach { data ->
                val label = data.date.dayOfWeek
                    .getDisplayName(TextStyle.NARROW, Locale.KOREAN)
                Text(
                    text = label,
                    style = MaterialTheme.typography.labelSmall,
                    color = if (data.date == today) Primary else TextSecondary,
                    textAlign = TextAlign.Center,
                    modifier = Modifier.weight(1f)
                )
            }
        }

        Spacer(Modifier.height(2.dp))

        // Step counts
        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.SpaceEvenly
        ) {
            weeklyData.forEach { data ->
                val label = if (data.steps >= 1000) "${data.steps / 1000}K" else "${data.steps}"
                Text(
                    text = label,
                    style = MaterialTheme.typography.labelSmall,
                    color = if (data.date == today) Primary else TextSecondary,
                    textAlign = TextAlign.Center,
                    modifier = Modifier.weight(1f)
                )
            }
        }
    }
}

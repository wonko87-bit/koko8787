package com.healthmonitor.ui.components

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.DirectionsWalk
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
import com.healthmonitor.ui.theme.StepsColor
import com.healthmonitor.ui.theme.StepsDim
import com.healthmonitor.ui.theme.TextSecondary
import java.text.NumberFormat

@Composable
fun StepsWidget(data: HealthData, modifier: Modifier = Modifier) {
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
                        colors = listOf(StepsDim.copy(alpha = 0.3f), CardSurface)
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
                        Icons.Default.DirectionsWalk,
                        contentDescription = null,
                        tint = StepsColor,
                        modifier = Modifier.size(18.dp)
                    )
                    Text(
                        text = "걸음수",
                        style = MaterialTheme.typography.titleMedium,
                        color = TextSecondary
                    )
                }

                Text(
                    text = NumberFormat.getNumberInstance().format(data.steps),
                    fontSize = 28.sp,
                    fontWeight = FontWeight.Bold,
                    color = StepsColor
                )

                Text(
                    text = "목표 ${NumberFormat.getNumberInstance().format(HealthData.STEP_GOAL)}",
                    style = MaterialTheme.typography.labelSmall,
                    color = TextSecondary
                )

                LinearProgressIndicator(
                    progress = { data.stepsProgressFraction },
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(6.dp)
                        .clip(RoundedCornerShape(3.dp)),
                    color = StepsColor,
                    trackColor = StepsDim
                )

                Text(
                    text = "${(data.stepsProgressFraction * 100).toInt()}%",
                    style = MaterialTheme.typography.labelSmall,
                    color = StepsColor
                )
            }
        }
    }
}

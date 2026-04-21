package com.healthmonitor.ui.components

import androidx.compose.foundation.background
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.WaterDrop
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
import com.healthmonitor.ui.theme.HydrationColor
import com.healthmonitor.ui.theme.HydrationDim
import com.healthmonitor.ui.theme.TextSecondary

@Composable
fun HydrationWidget(data: HealthData, modifier: Modifier = Modifier) {
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
                        colors = listOf(HydrationDim.copy(alpha = 0.3f), CardSurface)
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
                        Icons.Default.WaterDrop,
                        contentDescription = null,
                        tint = HydrationColor,
                        modifier = Modifier.size(18.dp)
                    )
                    Text(
                        text = "수분 섭취",
                        style = MaterialTheme.typography.titleMedium,
                        color = TextSecondary
                    )
                }

                Row(
                    verticalAlignment = Alignment.Bottom,
                    horizontalArrangement = Arrangement.spacedBy(2.dp)
                ) {
                    Text(
                        text = String.format("%.1f", data.hydrationLiters),
                        fontSize = 28.sp,
                        fontWeight = FontWeight.Bold,
                        color = HydrationColor
                    )
                    Text(
                        text = "L",
                        style = MaterialTheme.typography.bodyMedium,
                        color = TextSecondary
                    )
                }

                Text(
                    text = "목표 ${HealthData.HYDRATION_GOAL_LITERS}L",
                    style = MaterialTheme.typography.labelSmall,
                    color = TextSecondary
                )

                LinearProgressIndicator(
                    progress = { data.hydrationProgressFraction },
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(6.dp)
                        .clip(RoundedCornerShape(3.dp)),
                    color = HydrationColor,
                    trackColor = HydrationDim
                )

                Text(
                    text = "${(data.hydrationProgressFraction * 100).toInt()}%",
                    style = MaterialTheme.typography.labelSmall,
                    color = HydrationColor
                )
            }
        }
    }
}

package com.healthmonitor.ui.dashboard

import androidx.compose.foundation.layout.*
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Refresh
import androidx.compose.material.icons.filled.Warning
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.health.connect.client.permission.HealthPermission
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import androidx.lifecycle.viewmodel.compose.viewModel
import com.healthmonitor.data.health.HealthConnectManager
import com.healthmonitor.domain.models.HealthData
import com.healthmonitor.ui.components.*
import com.healthmonitor.ui.theme.CardSurface
import com.healthmonitor.ui.theme.Primary
import com.healthmonitor.ui.theme.TextSecondary
import kotlinx.coroutines.delay

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun DashboardScreen(viewModel: DashboardViewModel = androidx.hilt.navigation.compose.hiltViewModel()) {
    val uiState by viewModel.uiState.collectAsStateWithLifecycle()

    // Permission launcher for Health Connect
    val permissionLauncher = androidx.activity.compose.rememberLauncherForActivityResult(
        contract = androidx.health.connect.client.permission.PermissionController
            .createRequestPermissionResultContract()
    ) { granted ->
        viewModel.onPermissionsResult(granted)
    }

    // Auto-refresh every 5 minutes while the screen is visible
    LaunchedEffect(uiState.hasPermissions) {
        if (uiState.hasPermissions) {
            while (true) {
                delay(5 * 60 * 1000L)
                viewModel.syncData()
            }
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text("Health Monitor") },
                actions = {
                    if (uiState.lastSyncTime.isNotEmpty()) {
                        Text(
                            text = "동기화 ${uiState.lastSyncTime}",
                            style = MaterialTheme.typography.labelSmall,
                            color = TextSecondary,
                            modifier = Modifier.padding(end = 8.dp)
                        )
                    }
                    IconButton(
                        onClick = { viewModel.syncData() },
                        enabled = uiState.hasPermissions && !uiState.isLoading
                    ) {
                        Icon(Icons.Default.Refresh, contentDescription = "새로고침", tint = Primary)
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = MaterialTheme.colorScheme.background
                )
            )
        }
    ) { paddingValues ->
        Box(
            modifier = Modifier
                .fillMaxSize()
                .padding(paddingValues)
        ) {
            when {
                !uiState.isHealthConnectAvailable -> {
                    ErrorScreen(
                        icon = { Icon(Icons.Default.Warning, null, tint = MaterialTheme.colorScheme.error, modifier = Modifier.size(48.dp)) },
                        title = if (uiState.needsProviderUpdate) "Health Connect 업데이트 필요" else "Health Connect 미지원",
                        body = if (uiState.needsProviderUpdate)
                            "Play 스토어에서 Health Connect 앱을 업데이트해주세요."
                        else
                            "이 기기 또는 Android 버전에서는 Health Connect를 사용할 수 없습니다."
                    )
                }

                !uiState.hasPermissions -> {
                    ErrorScreen(
                        title = "권한이 필요합니다",
                        body = "걸음수, 수면, 칼로리, 수분 데이터를 읽으려면 Health Connect 권한을 허용해주세요.",
                        action = {
                            Button(
                                onClick = {
                                    permissionLauncher.launch(HealthConnectManager.REQUIRED_PERMISSIONS)
                                }
                            ) {
                                Text("권한 허용하기")
                            }
                        }
                    )
                }

                uiState.isLoading && uiState.todayData == null -> {
                    Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                        CircularProgressIndicator(color = Primary)
                    }
                }

                else -> {
                    DashboardContent(
                        uiState = uiState,
                        modifier = Modifier.fillMaxSize()
                    )
                }
            }
        }
    }
}

@Composable
private fun DashboardContent(uiState: DashboardUiState, modifier: Modifier = Modifier) {
    val today = uiState.todayData ?: HealthData(
        date = java.time.LocalDate.now(),
        steps = 0, sleepMinutes = 0, activeCaloriesKcal = 0.0, hydrationLiters = 0.0
    )

    Column(
        modifier = modifier
            .verticalScroll(rememberScrollState())
            .padding(horizontal = 16.dp, vertical = 8.dp),
        verticalArrangement = Arrangement.spacedBy(16.dp)
    ) {
        if (uiState.error != null) {
            Surface(
                color = MaterialTheme.colorScheme.errorContainer,
                shape = RoundedCornerShape(12.dp)
            ) {
                Text(
                    text = uiState.error,
                    modifier = Modifier.padding(12.dp),
                    style = MaterialTheme.typography.bodyMedium,
                    color = MaterialTheme.colorScheme.onErrorContainer
                )
            }
        }

        // Today's stats grid (2x2)
        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.spacedBy(12.dp)
        ) {
            StepsWidget(data = today, modifier = Modifier.weight(1f).height(160.dp))
            SleepWidget(data = today, modifier = Modifier.weight(1f).height(160.dp))
        }

        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.spacedBy(12.dp)
        ) {
            CaloriesWidget(data = today, modifier = Modifier.weight(1f).height(160.dp))
            HydrationWidget(data = today, modifier = Modifier.weight(1f).height(160.dp))
        }

        // Weekly activity chart
        if (uiState.weeklyData.isNotEmpty()) {
            Card(
                modifier = Modifier.fillMaxWidth(),
                shape = RoundedCornerShape(16.dp),
                colors = CardDefaults.cardColors(containerColor = CardSurface)
            ) {
                Column(modifier = Modifier.padding(16.dp)) {
                    Text(
                        text = "주간 활동 (걸음수 · 최근 7일)",
                        style = MaterialTheme.typography.titleMedium,
                        color = TextSecondary
                    )
                    Spacer(Modifier.height(12.dp))
                    WeeklyActivityChart(
                        weeklyData = uiState.weeklyData,
                        modifier = Modifier.fillMaxWidth()
                    )
                }
            }
        }

        Spacer(Modifier.height(8.dp))
    }
}

@Composable
private fun ErrorScreen(
    title: String,
    body: String,
    icon: (@Composable () -> Unit)? = null,
    action: (@Composable () -> Unit)? = null
) {
    Box(Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
        Column(
            horizontalAlignment = Alignment.CenterHorizontally,
            verticalArrangement = Arrangement.spacedBy(16.dp),
            modifier = Modifier.padding(32.dp)
        ) {
            icon?.invoke()
            Text(
                text = title,
                style = MaterialTheme.typography.titleLarge,
                textAlign = TextAlign.Center
            )
            Text(
                text = body,
                style = MaterialTheme.typography.bodyMedium,
                color = TextSecondary,
                textAlign = TextAlign.Center
            )
            action?.invoke()
        }
    }
}

package com.healthmonitor.ui.dashboard

import android.content.Context
import androidx.health.connect.client.HealthConnectClient
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import com.healthmonitor.data.health.HealthConnectManager
import com.healthmonitor.data.repository.HealthRepository
import com.healthmonitor.domain.models.HealthData
import com.healthmonitor.worker.HealthSyncWorker
import dagger.hilt.android.lifecycle.HiltViewModel
import dagger.hilt.android.qualifiers.ApplicationContext
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.update
import kotlinx.coroutines.launch
import java.time.LocalTime
import java.time.format.DateTimeFormatter
import javax.inject.Inject

data class DashboardUiState(
    val isLoading: Boolean = true,
    val sdkStatus: Int = HealthConnectClient.SDK_UNAVAILABLE,
    val hasPermissions: Boolean = false,
    val todayData: HealthData? = null,
    val weeklyData: List<HealthData> = emptyList(),
    val lastSyncTime: String = "",
    val error: String? = null
) {
    val isHealthConnectAvailable: Boolean
        get() = sdkStatus == HealthConnectClient.SDK_AVAILABLE
    val needsProviderUpdate: Boolean
        get() = sdkStatus == HealthConnectClient.SDK_UNAVAILABLE_PROVIDER_UPDATE_REQUIRED
}

@HiltViewModel
class DashboardViewModel @Inject constructor(
    @ApplicationContext private val context: Context,
    private val repository: HealthRepository
) : ViewModel() {

    private val _uiState = MutableStateFlow(DashboardUiState())
    val uiState: StateFlow<DashboardUiState> = _uiState.asStateFlow()

    init {
        checkAvailabilityAndPermissions()
    }

    fun checkAvailabilityAndPermissions() {
        viewModelScope.launch {
            val status = HealthConnectClient.getSdkStatus(context)
            _uiState.update { it.copy(sdkStatus = status) }

            if (status == HealthConnectClient.SDK_AVAILABLE) {
                val hasPermissions = repository.hasAllPermissions()
                _uiState.update { it.copy(hasPermissions = hasPermissions) }
                if (hasPermissions) {
                    syncData()
                    HealthSyncWorker.schedule(context)
                } else {
                    _uiState.update { it.copy(isLoading = false) }
                }
            } else {
                _uiState.update { it.copy(isLoading = false) }
            }
        }
    }

    fun onPermissionsResult(granted: Set<String>) {
        val hasAll = granted.containsAll(HealthConnectManager.REQUIRED_PERMISSIONS)
        _uiState.update { it.copy(hasPermissions = hasAll) }
        if (hasAll) {
            syncData()
            HealthSyncWorker.schedule(context)
        }
    }

    fun syncData() {
        viewModelScope.launch {
            _uiState.update { it.copy(isLoading = true, error = null) }
            runCatching { repository.syncWeeklyData() }
                .onSuccess { weekly ->
                    _uiState.update {
                        it.copy(
                            isLoading = false,
                            weeklyData = weekly,
                            todayData = weekly.lastOrNull(),
                            lastSyncTime = LocalTime.now()
                                .format(DateTimeFormatter.ofPattern("HH:mm"))
                        )
                    }
                }
                .onFailure { e ->
                    _uiState.update {
                        it.copy(isLoading = false, error = "동기화 실패: ${e.message}")
                    }
                }
        }
    }
}

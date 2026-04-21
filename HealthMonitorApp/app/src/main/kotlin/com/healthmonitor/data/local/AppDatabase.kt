package com.healthmonitor.data.local

import androidx.room.Database
import androidx.room.RoomDatabase
import com.healthmonitor.data.local.dao.HealthDataDao
import com.healthmonitor.data.local.entities.HealthDataEntity

@Database(entities = [HealthDataEntity::class], version = 1, exportSchema = false)
abstract class AppDatabase : RoomDatabase() {
    abstract fun healthDataDao(): HealthDataDao
}

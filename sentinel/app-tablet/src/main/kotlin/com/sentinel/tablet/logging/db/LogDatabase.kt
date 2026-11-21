package com.sentinel.tablet.logging.db

import android.content.Context
import androidx.room.Database
import androidx.room.Room
import androidx.room.RoomDatabase

@Database(
    entities = [SessionRecord::class, EventRecord::class],
    version = 1,
    exportSchema = false,
)
abstract class LogDatabase : RoomDatabase() {
    abstract fun logDao(): LogDao

    companion object {
        @Volatile
        private var instance: LogDatabase? = null

        fun get(context: Context): LogDatabase =
            instance ?: synchronized(this) {
                instance ?: Room.databaseBuilder(
                    context.applicationContext,
                    LogDatabase::class.java,
                    "sentinel_logs.db",
                ).build().also { instance = it }
            }
    }
}

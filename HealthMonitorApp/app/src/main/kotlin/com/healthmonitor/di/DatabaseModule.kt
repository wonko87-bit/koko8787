package com.healthmonitor.di

import android.content.Context
import androidx.room.Room
import com.healthmonitor.data.local.AppDatabase
import com.healthmonitor.data.local.dao.HealthDataDao
import dagger.Module
import dagger.Provides
import dagger.hilt.InstallIn
import dagger.hilt.android.qualifiers.ApplicationContext
import dagger.hilt.components.SingletonComponent
import javax.inject.Singleton

@Module
@InstallIn(SingletonComponent::class)
object DatabaseModule {

    @Provides
    @Singleton
    fun provideDatabase(@ApplicationContext context: Context): AppDatabase =
        Room.databaseBuilder(context, AppDatabase::class.java, "health_monitor_db").build()

    @Provides
    fun provideHealthDataDao(db: AppDatabase): HealthDataDao = db.healthDataDao()
}

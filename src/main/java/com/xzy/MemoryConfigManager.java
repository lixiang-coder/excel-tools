package com.xzy;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * 内存配置管理器
 * 根据系统内存动态调整处理参数，解决低配置机器OOM问题
 */
public class MemoryConfigManager {
    private static final Logger logger = LogManager.getLogger(MemoryConfigManager.class);
    
    // 内存阈值配置
    private static final long LOW_MEMORY_THRESHOLD = 512;   // 512MB
    private static final long MEDIUM_MEMORY_THRESHOLD = 1024; // 1GB
    private static final double MEMORY_WARNING_THRESHOLD = 0.85; // 85%内存使用率警告
    
    // 配置参数
    private final int memoryWindowSize;
    private final int batchSize;
    private final int maxFileSize;
    private final boolean enableCompression;
    
    private static MemoryConfigManager instance;
    
    private MemoryConfigManager() {
        long maxMemoryMB = getMaxMemoryMB();
        logger.info("系统最大内存: {}MB", maxMemoryMB);
        
        // 根据内存大小动态调整参数
        if (maxMemoryMB < LOW_MEMORY_THRESHOLD) {
            // 低内存配置 (<512MB)
            this.memoryWindowSize = 5;
            this.batchSize = 50;
            this.maxFileSize = 50; // 50MB
            this.enableCompression = true;
            logger.info("应用低内存配置");
        } else if (maxMemoryMB < MEDIUM_MEMORY_THRESHOLD) {
            // 中等内存配置 (512MB-1GB)
            this.memoryWindowSize = 20;
            this.batchSize = 200;
            this.maxFileSize = 200; // 200MB
            this.enableCompression = true;
            logger.info("应用中等内存配置");
        } else {
            // 高内存配置 (>1GB)
            this.memoryWindowSize = 50;
            this.batchSize = 500;
            this.maxFileSize = 500; // 500MB
            this.enableCompression = false;
            logger.info("应用高内存配置");
        }
        
        logger.info("内存配置 - 窗口大小: {}, 批处理大小: {}, 最大文件: {}MB", 
                   memoryWindowSize, batchSize, maxFileSize);
    }
    
    public static synchronized MemoryConfigManager getInstance() {
        if (instance == null) {
            instance = new MemoryConfigManager();
        }
        return instance;
    }
    
    /**
     * 获取系统最大内存(MB)
     */
    private long getMaxMemoryMB() {
        return Runtime.getRuntime().maxMemory() / 1024 / 1024;
    }
    
    /**
     * 获取当前内存使用率
     */
    public double getCurrentMemoryUsage() {
        Runtime runtime = Runtime.getRuntime();
        long maxMemory = runtime.maxMemory();
        long totalMemory = runtime.totalMemory();
        long freeMemory = runtime.freeMemory();
        long usedMemory = totalMemory - freeMemory;
        
        return (double) usedMemory / maxMemory;
    }
    
    /**
     * 检查内存使用是否超过警告阈值
     */
    public boolean checkMemoryWarning() {
        double usage = getCurrentMemoryUsage();
        if (usage > MEMORY_WARNING_THRESHOLD) {
            logger.warn("内存使用率过高: {:.2f}%", usage * 100);
            return true;
        }
        return false;
    }
    
    /**
     * 强制垃圾回收并等待
     */
    public void forceGarbageCollection() {
        long beforeGC = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        System.gc();
        try {
            Thread.sleep(100); // 给GC时间执行
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
        long afterGC = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        logger.debug("GC回收内存: {}MB", (beforeGC - afterGC) / 1024 / 1024);
    }
    
    /**
     * 获取内存使用情况描述
     */
    public String getMemoryStatus() {
        Runtime runtime = Runtime.getRuntime();
        long maxMemory = runtime.maxMemory() / 1024 / 1024;
        long totalMemory = runtime.totalMemory() / 1024 / 1024;
        long freeMemory = runtime.freeMemory() / 1024 / 1024;
        long usedMemory = totalMemory - freeMemory;
        
        return String.format("内存使用: %dMB/%dMB (%.1f%%)", 
                           usedMemory, maxMemory, getCurrentMemoryUsage() * 100);
    }
    
    // Getter方法
    public int getMemoryWindowSize() { return memoryWindowSize; }
    public int getBatchSize() { return batchSize; }
    public int getMaxFileSize() { return maxFileSize; }
    public boolean isEnableCompression() { return enableCompression; }
    
    /**
     * 根据文件大小调整批处理大小
     */
    public int getOptimalBatchSize(long fileSizeMB) {
        if (fileSizeMB > 100) {
            return Math.max(batchSize / 4, 10); // 大文件使用更小的批次
        } else if (fileSizeMB > 50) {
            return Math.max(batchSize / 2, 20);
        }
        return batchSize;
    }
    
    /**
     * 检查文件是否可以处理
     */
    public boolean canProcessFile(long fileSizeMB) {
        return fileSizeMB <= maxFileSize;
    }
}
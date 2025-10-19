@echo off
echo 启动Excel工具集 - 内存优化版本
echo ================================

REM 获取系统内存信息
for /f "skip=1" %%i in ('wmic OS get TotalVisibleMemorySize /value') do if not "%%i"=="" set %%i
set /a MEMORY_MB=%TotalVisibleMemorySize%/1024

echo 系统总内存: %MEMORY_MB%MB

REM 根据系统内存设置JVM参数
if %MEMORY_MB% LSS 1024 (
    echo 检测到低内存配置 - 应用保守设置
    set XMS=128m
    set XMX=512m
    set GC_PARAMS=-XX:+UseSerialGC
) else if %MEMORY_MB% LSS 2048 (
    echo 检测到中等内存配置 - 应用平衡设置
    set XMS=256m
    set XMX=1024m
    set GC_PARAMS=-XX:+UseG1GC -XX:MaxGCPauseMillis=200
) else if %MEMORY_MB% LSS 4096 (
    echo 检测到较高内存配置 - 应用增强设置
    set XMS=512m
    set XMX=2560m
    set GC_PARAMS=-XX:+UseG1GC -XX:MaxGCPauseMillis=100 -XX:+UseStringDeduplication
) else if %MEMORY_MB% LSS 8192 (
    echo 检测到高内存配置 - 应用高性能设置（支持大文件）
    set XMS=1024m
    set XMX=6144m
    set GC_PARAMS=-XX:+UseG1GC -XX:MaxGCPauseMillis=100 -XX:+UseStringDeduplication -XX:G1HeapRegionSize=16m
) else (
    echo 检测到超高内存配置 - 应用极致性能设置（支持超大文件）
    set XMS=2048m
    set XMX=8192m
    set GC_PARAMS=-XX:+UseG1GC -XX:MaxGCPauseMillis=100 -XX:+UseStringDeduplication -XX:G1HeapRegionSize=16m -XX:+ParallelRefProcEnabled
)

echo 使用JVM参数: -Xms%XMS% -Xmx%XMX%
echo 启动程序...
echo.

java -Xms%XMS% -Xmx%XMX% %GC_PARAMS% ^
     -Dfile.encoding=UTF-8 ^
     -Dsun.java2d.noddraw=true ^
     -Djava.awt.headless=false ^
     -jar target\excel-processor-1.0-SNAPSHOT.jar

if errorlevel 1 (
    echo.
    echo 程序异常退出，可能的原因：
    echo 1. 内存不足 - 请关闭其他程序后重试，或手动指定更大内存：
    echo    java -Xmx4g -jar target\excel-processor-1.0-SNAPSHOT.jar
    echo 2. Java版本不兼容 - 需要Java 8或更高版本
    echo 3. 文件损坏 - 请重新编译项目
    echo 4. Excel文件过于复杂 - 建议将大文件拆分为多个小文件
    echo.
    pause
) else (
    echo 程序正常退出
)
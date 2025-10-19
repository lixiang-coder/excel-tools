@echo off
chcp 65001 > nul
title Excel工具集 - 启动器
color 0A

cls
echo ╔════════════════════════════════════════════════╗
echo ║                                                ║
echo ║          Excel工具集 - 智能启动器              ║
echo ║                                                ║
echo ╚════════════════════════════════════════════════╝
echo.
echo.
echo 请选择启动方式：
echo.
echo   [1] 快速启动（小文件模式，<100MB）
echo       └─ 内存分配：默认（约512MB-1GB）
echo       └─ 适用场景：日常使用，小文件处理
echo.
echo   [2] 标准启动（中等文件模式，100-200MB）
echo       └─ 内存分配：智能检测（根据系统自动分配）
echo       └─ 适用场景：推荐使用，自动适配
echo.
echo   [3] 高性能启动（大文件模式，200MB+）
echo       └─ 内存分配：6GB
echo       └─ 适用场景：处理超大文件，如250MB+
echo.
echo   [4] 极致性能启动（超大文件模式，500MB+）
echo       └─ 内存分配：8GB
echo       └─ 适用场景：处理超大复杂文件
echo.
echo   [0] 退出
echo.
echo ════════════════════════════════════════════════
echo.

set /p choice=请输入选项 [0-4]:

if "%choice%"=="1" goto quick
if "%choice%"=="2" goto optimized
if "%choice%"=="3" goto large
if "%choice%"=="4" goto ultra
if "%choice%"=="0" goto end
echo 无效选项，请重新选择
pause
goto :eof

:quick
cls
echo [快速启动模式]
echo 启动中...
echo.
java -jar target\excel-processor-1.0-SNAPSHOT.jar
goto end

:optimized
cls
echo [标准启动模式 - 智能内存分配]
echo 正在检测系统内存...
echo.
call start_optimized.bat
goto end

:large
cls
echo [高性能启动模式 - 6GB内存]
echo 启动中...
echo.
java -Xms1024m -Xmx6144m ^
     -XX:+UseG1GC ^
     -XX:MaxGCPauseMillis=100 ^
     -XX:+UseStringDeduplication ^
     -Dfile.encoding=UTF-8 ^
     -jar target\excel-processor-1.0-SNAPSHOT.jar
goto end

:ultra
cls
echo [极致性能启动模式 - 8GB内存]
echo 启动中...
echo.
java -Xms2048m -Xmx8192m ^
     -XX:+UseG1GC ^
     -XX:MaxGCPauseMillis=100 ^
     -XX:+UseStringDeduplication ^
     -XX:G1HeapRegionSize=16m ^
     -Dfile.encoding=UTF-8 ^
     -jar target\excel-processor-1.0-SNAPSHOT.jar
goto end

:end
exit

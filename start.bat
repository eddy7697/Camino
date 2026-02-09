@echo off
chcp 65001 >nul
title 朝聖之路 Camino de Santiago

echo ============================================
echo   朝聖之路 - 啟動網頁伺服器
echo ============================================
echo.

:: Check if Docker is running
docker info >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Docker 尚未啟動，請先開啟 Docker Desktop
    echo     啟動後再重新執行此腳本
    pause
    exit /b 1
)

:: Stop old container if exists
docker rm -f camino 2>nul

:: Build image
echo [1/3] 建置 Docker 映像檔...
docker build -t camino-web "%~dp0"
if %errorlevel% neq 0 (
    echo [!] 建置失敗
    pause
    exit /b 1
)

:: Run container
echo [2/3] 啟動容器...
docker run -d --name camino -p 8080:8080 camino-web
if %errorlevel% neq 0 (
    echo [!] 啟動失敗
    pause
    exit /b 1
)

:: Open browser
echo [3/3] 開啟瀏覽器...
timeout /t 1 >nul
start http://localhost:8080

echo.
echo ============================================
echo   網頁已啟動：http://localhost:8080
echo   按任意鍵停止伺服器並關閉...
echo ============================================
pause >nul

:: Cleanup
docker stop camino >nul 2>&1
docker rm camino >nul 2>&1
echo 伺服器已停止。

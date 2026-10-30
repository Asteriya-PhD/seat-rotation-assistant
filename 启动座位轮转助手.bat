@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "SeatRotator.exe" (
  echo 未找到 SeatRotator.exe，请确认压缩包已完整解压。
  echo.
  pause
  exit /b 1
)

start "" "SeatRotator.exe"

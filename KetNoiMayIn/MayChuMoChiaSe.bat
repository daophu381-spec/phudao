@echo off
rem ============================================================
rem  CHAY TREN MAY DANG CAM MAY IN (may chia se, vd 10.5.208.100)
rem  Bam dup de mo khoa chia se may in cho cac may khac.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Dang xin quyen Administrator...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0MayChuMoChiaSe.ps1"
echo.
pause

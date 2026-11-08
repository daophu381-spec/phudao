@echo off
rem ============================================================
rem  BAM DUP CHUOT FILE NAY DE CAI DRIVER HP LaserJet Pro M402
rem  (tu xin quyen Admin, tu nhan Windows 32/64 bit)
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Dang xin quyen Administrator...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0CaiDriverM402.ps1" %*
echo.
pause

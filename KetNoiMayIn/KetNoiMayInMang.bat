@echo off
rem ============================================================
rem  CHAY TREN MAY CAN IN: ket noi may in chia se (mac dinh 10.5.208.100)
rem  Dung: bam dup. Hoac:  KetNoiMayInMang.bat 10.5.208.100
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Dang xin quyen Administrator...
    if "%~1"=="" (
        powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    ) else (
        powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '%~1' -Verb RunAs"
    )
    exit /b
)

if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0KetNoiMayInMang.ps1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0KetNoiMayInMang.ps1" -IP "%~1"
)
echo.
pause

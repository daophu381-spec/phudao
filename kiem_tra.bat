@echo off
rem Kéo thả file .docx lên file này: chương trình kiểm tra rồi mở báo cáo trong trình duyệt.
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" (
  echo Cach dung: keo tha file .docx vao file kiem_tra.bat
  pause
  exit /b
)
python -m kiemtra_vanban "%~1" --html "%~dpn1_kiemtra.html"
start "" "%~dpn1_kiemtra.html"
pause

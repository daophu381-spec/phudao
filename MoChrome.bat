@echo off
rem Mo Chrome rieng cho cong cu nhap tu dong (dang nhap 1 lan, lan sau nho san).
rem Hay DONG HET Chrome dang mo truoc khi chay file nay.
set CHROME="C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist %CHROME% set CHROME="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist %CHROME% set CHROME="%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
start "" %CHROME% --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\ChromeNhapTuDong"

@echo off
chcp 65001 >nul 2>nul
cd /d "%~dp0"

echo [%time%] Python 프로세스 종료 중...
taskkill /F /IM python.exe  >nul 2>nul
taskkill /F /IM pythonw.exe >nul 2>nul
taskkill /F /IM pyw.exe     >nul 2>nul
taskkill /F /IM py.exe      >nul 2>nul
timeout /t 2 /nobreak >nul

echo [%time%] 창 없이 서버 재시작 중...
call run_server.bat

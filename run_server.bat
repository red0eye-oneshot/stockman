@echo off
chcp 65001 >nul 2>nul
setlocal
cd /d "%~dp0"

set "LOGFILE=%~dp0py_start.log"
echo [%date% %time%] run_server.bat started > "%LOGFILE%"

rem Python 실행 파일 탐색 (py 우선)
set "PYEXE="
where py >nul 2>nul
if not errorlevel 1 (
    py -c "import requests" >nul 2>nul
    if not errorlevel 1 set "PYEXE=py"
)
if not defined PYEXE (
    for /f "tokens=*" %%P in ('where python 2^>nul') do (
        if not defined PYEXE (
            "%%P" -c "import requests" >nul 2>nul
            if not errorlevel 1 set "PYEXE=%%P"
        )
    )
)
echo [%time%] PYEXE=%PYEXE% >> "%LOGFILE%"
if not defined PYEXE ( echo [ERROR] Python not found >> "%LOGFILE%" & endlocal & exit /b 1 )
if not exist "%~dp0stock_server.py" ( echo [ERROR] stock_server.py missing >> "%LOGFILE%" & endlocal & exit /b 1 )

rem 이미 서버 실행 중이면 브라우저만 열고 종료
curl -s --max-time 2 http://127.0.0.1:5555/api/ping >nul 2>nul
if not errorlevel 1 (
    echo [%time%] Server already running >> "%LOGFILE%"
    start "" "http://127.0.0.1:5555"
    endlocal & exit /b 0
)

rem 포트 사용 중이면 기존 프로세스 정리
netstat -an | findstr ":5555 " >nul 2>nul
if not errorlevel 1 (
    echo [%time%] Port 5555 busy - killing old process >> "%LOGFILE%"
    taskkill /F /IM python.exe >nul 2>nul
    taskkill /F /IM pythonw.exe >nul 2>nul
    taskkill /F /IM py.exe >nul 2>nul
    timeout /t 2 /nobreak >nul
)

set "SRVPATH=%~dp0stock_server.py"
echo [%time%] Starting server (background)... >> "%LOGFILE%"

rem pythonw 있으면 창 없이, 없으면 최소화 창으로 실행
set "PYEXEW="
if /I "%PYEXE%"=="py" (
    where pyw >nul 2>nul
    if not errorlevel 1 set "PYEXEW=pyw"
) else (
    set "PYEXEW=%PYEXE:python.exe=pythonw.exe%"
    if not exist "%PYEXEW%" set "PYEXEW="
)

if defined PYEXEW (
    echo [%time%] Using %PYEXEW% ^(no window^) >> "%LOGFILE%"
    start "" "%PYEXEW%" "%SRVPATH%"
) else (
    rem pyw/pythonw 없을 때 VBScript으로 완전 숨김 실행
    echo [%time%] Using VBScript hidden launch >> "%LOGFILE%"
    set "VBSPATH=%TEMP%\run_stock_hidden.vbs"
    echo Set oShell = CreateObject^("WScript.Shell"^) > "%VBSPATH%"
    echo oShell.Run """%PYEXE%"" ""%SRVPATH%""", 0, False >> "%VBSPATH%"
    wscript.exe "%VBSPATH%"
)

rem 서버 올라올 때까지 최대 40초 대기
set /a "waited=0"
:waitloop
timeout /t 1 /nobreak >nul
set /a "waited+=1"
curl -s --max-time 1 http://127.0.0.1:5555/api/ping >nul 2>nul
if not errorlevel 1 goto :open
echo [%time%] waiting... %waited%s >> "%LOGFILE%"
if %waited% geq 40 (
    echo [ERROR] Timeout - server did not start within 40s >> "%LOGFILE%"
    endlocal & exit /b 1
)
goto :waitloop

:open
echo [%time%] Server UP after %waited%s >> "%LOGFILE%"
start "" "http://127.0.0.1:5555"
endlocal & exit /b 0

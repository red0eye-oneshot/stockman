@echo off
cd /d "%~dp0"
echo === Stock Server Diagnose ===
echo.

where py >nul 2>nul
if not errorlevel 1 (
    py -c "import requests" >nul 2>nul
    if not errorlevel 1 (
        echo [OK] Python=py
        goto :found
    )
)

for /f "tokens=*" %%P in ('where python 2^>nul') do (
    if not defined PYEXE (
        "%%P" -c "import requests" >nul 2>nul
        if not errorlevel 1 set "PYEXE=%%P"
    )
)
if defined PYEXE (
    echo [OK] Python=%PYEXE%
    goto :found
)

echo [FAIL] Python or requests not found
echo Run: pip install requests
pause
exit /b 1

:found
echo [OK] stock_server.py exists? 
if exist "%~dp0stock_server.py" (echo YES) else (echo NO - file missing!)
echo.
echo [TEST] Checking port 5555...
curl -s --max-time 2 http://localhost:5555/api/ping
if not errorlevel 1 (
    echo.
    echo [INFO] Server already running - do Ctrl+Shift+R in browser
    pause
    exit /b 0
)
echo.
echo [START] Starting server now... errors will show below
echo.
if defined PYEXE (
    "%PYEXE%" "%~dp0stock_server.py"
) else (
    py "%~dp0stock_server.py"
)
echo.
echo [EXIT] Server stopped
pause
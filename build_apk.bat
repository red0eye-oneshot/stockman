@echo off
setlocal enabledelayedexpansion

set SITE_URL=https://stockman-10k4.onrender.com
set APK_OUT=%~dp0stock-tracker.apk
set TWA_DIR=%~dp0twa-build
set NODE_MSI=%TEMP%\node_setup.msi
set NODE_DL=https://nodejs.org/dist/v20.19.2/node-v20.19.2-x64.msi

echo.
echo === Stock Tracker APK Builder ===
echo.

:: Step 1: Wake up Render server (free tier sleeps)
echo [1] Waking up Render server... (wait up to 90 sec)
curl -s -o nul --max-time 90 --retry 5 --retry-delay 10 --retry-connrefused "%SITE_URL%/api/ping"
echo  Server contacted. Waiting 5 sec...
timeout /t 5 /nobreak >nul
echo.

:: Step 2: Node.js
where node >nul 2>&1
if errorlevel 1 goto :install_node
for /f "tokens=*" %%v in ('node -v') do echo [2] Node.js %%v found
goto :check_bw

:install_node
echo [2] Downloading Node.js (~100MB)...
curl -L -o "%NODE_MSI%" "%NODE_DL%"
if errorlevel 1 ( echo ERROR: Download failed. & goto :end )
echo [2] Installing Node.js...
msiexec /i "%NODE_MSI%" /qn /norestart
set "PATH=%PATH%;C:\Program Files\nodejs"
echo [2] Node.js installed

:check_bw
:: Step 3: Bubblewrap
where bubblewrap >nul 2>&1
if errorlevel 1 goto :install_bw
echo [3] Bubblewrap found
goto :init

:install_bw
echo [3] Installing Bubblewrap...
call npm install -g @bubblewrap/cli
if errorlevel 1 ( echo ERROR: Bubblewrap install failed. & goto :end )
echo [3] Bubblewrap installed

:init
:: Step 4: Init (first run only)
if not exist "%TWA_DIR%" mkdir "%TWA_DIR%"
cd /d "%TWA_DIR%"
if exist "twa-manifest.json" goto :build

echo.
echo [4] First run setup...
echo     Press ENTER for all defaults. Set a signing key password.
echo.
bubblewrap init --manifest "%SITE_URL%/manifest.json"
if errorlevel 1 (
    echo ERROR: Init failed.
    cd /d "%~dp0"
    rmdir /s /q "%TWA_DIR%" 2>nul
    goto :end
)
echo [4] Init done

:build
:: Step 5: Build
echo.
echo [5] Building APK... (2-5 min)
echo.
bubblewrap build
if errorlevel 1 (
    echo ERROR: Build failed.
    cd /d "%~dp0"
    goto :end
)

:: Copy APK
set FOUND=0
for /r . %%f in (*.apk) do (
    if !FOUND!==0 (
        copy "%%f" "%APK_OUT%" >nul
        set FOUND=1
    )
)
cd /d "%~dp0"

if !FOUND!==1 (
    echo.
    echo === BUILD SUCCESS: stock-tracker.apk ===
) else (
    echo WARNING: APK not found. Check twa-build folder.
)

:end
echo.
pause

@echo off
cd /d "%~dp0"

echo.
echo ================================================
echo   GitHub Upload - stock-tracker (stockman)
echo   Repo: red0eye-oneshot/stockman
echo ================================================
echo.
echo [Folder] %~dp0
echo.

echo [Step 1] Checking Git...
git --version
if errorlevel 1 (
    echo.
    echo [ERROR] Git not found!
    echo Install: https://git-scm.com/download/win
    pause
    exit /b 1
)
echo.

echo [Step 2] Git user config...
git config --global user.email "red0eye@gmail.com"
git config --global user.name "red0eye"
git config user.email
git config user.name
echo.

echo [Step 3] Check/init repo...
if not exist ".git" (
    echo  No .git folder - initializing...
    git init
    git remote add origin https://github.com/red0eye-oneshot/stockman.git
    echo  Done.
) else (
    echo  .git exists OK
    git remote -v
)
echo.

echo [Step 4] Changed files:
echo ----------------------------------------
git status --short
echo ----------------------------------------
echo.

set COMMIT_MSG=
set /p COMMIT_MSG="Commit message (Enter = auto date): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=Update %date%
echo Commit: %COMMIT_MSG%
echo.

echo [Step 5] git add...
git add -A
echo Done.
echo.

echo [Step 6] git commit...
git commit -m "%COMMIT_MSG%"
echo.

echo [Step 7] git push...
git push -u origin HEAD
echo.

if errorlevel 1 (
    echo [PUSH FAILED] Try manually:
    echo   git push -u origin main
    echo   git push -u origin master
    echo.
) else (
    echo ================================================
    echo   [SUCCESS] Upload complete!
    echo   https://github.com/red0eye-oneshot/stockman
    echo ================================================
    echo.
    echo Next: run build_apk.bat
)

echo.
pause

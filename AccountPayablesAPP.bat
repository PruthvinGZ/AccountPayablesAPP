@echo off
setlocal enabledelayedexpansion

:: Force admin rights (critical for Program Files access)
NET FILE >NUL 2>&1
IF %ERRORLEVEL% NEQ 0 (
    PowerShell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    EXIT /B
)

:: Set working directory to script location
cd /d "%~dp0"

:: Cleanup previous instances
taskkill /F /IM pythonw.exe >NUL 2>&1
timeout /t 3 >NUL

:: Start server with debug logging
echo Starting server...
start "AppServer" /min cmd /c ""%~dp0venv\Scripts\pythonw.exe" "%~dp0app.py" > "%~dp0logs\app_startup.log" 2>&1"

:: Wait for port file (extended timeout)
echo Waiting for server start...
set PORT=
set MAX_ATTEMPTS=30
set ATTEMPT=0

:check_port
timeout /t 2 >NUL
set /a ATTEMPT+=1

if exist "%~dp0logs\server_port.txt" (
    set /p PORT=<"%~dp0logs\server_port.txt"
    set "PORT=!PORT:"=!"
    set "PORT=!PORT: =!"
    echo Detected port: !PORT!
    goto verify_server
) else (
    echo Attempt !ATTEMPT!/!MAX_ATTEMPTS!: No port file yet
)

if !ATTEMPT! GEQ !MAX_ATTEMPTS! (
    echo Server failed to start after 60 seconds
    echo === Server Log ===
    type "%~dp0logs\app_startup.log"
    pause
    exit /b 1
)
goto check_port

:verify_server
echo Verifying server on port !PORT!...
curl -s -o nul http://127.0.0.1:!PORT!/health
if %ERRORLEVEL% neq 0 (
    echo Server not responding on port !PORT!
    goto check_port
)

:: Open browser
start "" "http://127.0.0.1:!PORT!/"
echo Server running on http://127.0.0.1:!PORT!/
exit /b 0
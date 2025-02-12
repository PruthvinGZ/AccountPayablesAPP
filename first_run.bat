@echo off
setlocal enabledelayedexpansion

:: Configure paths
set "APP_DIR=%~dp0"
cd /d "%APP_DIR%"

:: Verify Python installation
echo Checking Python installation...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: Python not found in PATH
    echo 1. Verify Python is installed
    echo 2. Check "Add Python to PATH" was selected during installation
    echo 3. Restart computer after installation
    pause
    exit /b 1
)

:: Create clean virtual environment
echo Creating virtual environment...
if exist "%APP_DIR%venv" (
    echo Removing existing virtual environment...
    rmdir /s /q "%APP_DIR%venv"
)

python -m venv "%APP_DIR%venv"
if %ERRORLEVEL% neq 0 (
    echo Failed to create virtual environment
    pause
    exit /b 1
)

:: Install requirements
echo Installing dependencies...
if not exist "requirements.txt" (
    echo Missing requirements.txt
    pause
    exit /b 1
)

call "%APP_DIR%venv\Scripts\pip.exe" install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo Failed to install requirements
    pause
    exit /b 1
)

:: Verify critical packages
echo Verifying installations...
"%APP_DIR%venv\Scripts\python.exe" -c "import flask, openpyxl" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Critical packages missing
    call "%APP_DIR%venv\Scripts\pip.exe" install flask openpyxl
)

echo Environment setup completed successfully
echo Run AccountPayablesAPP.bat to start the application
pause
exit /b 0
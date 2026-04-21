@echo off
setlocal enabledelayedexpansion

echo ========================================
echo CCTV IP Toolkit - Build Script
echo ========================================
echo.
echo This script will:
echo   1. Check for Python (install if missing)
echo   2. Install required packages
echo   3. Build the standalone .exe
echo.

REM -- Step 1: Check for Python --
echo [1/4] Checking for Python...
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found. Downloading Python installer...
    echo.

    set "PYTHON_URL=https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"
    set "PYTHON_INSTALLER=python-installer.exe"

    REM Try PowerShell download
    powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%'" 2>nul
    if not exist "%PYTHON_INSTALLER%" (
        curl -L -o "%PYTHON_INSTALLER%" "%PYTHON_URL%" 2>nul
    )
    if not exist "%PYTHON_INSTALLER%" (
        certutil -urlcache -split -f "%PYTHON_URL%" "%PYTHON_INSTALLER%" >nul 2>&1
    )

    if not exist "%PYTHON_INSTALLER%" (
        echo.
        echo ERROR: Could not download Python installer.
        echo Please install Python manually from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during installation.
        echo Then re-run this script.
        pause
        exit /b 1
    )

    echo Installing Python 3.12 ...
    echo   - Adding to PATH
    echo   - Installing pip
    echo.

    "%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0

    if %errorlevel% neq 0 (
        echo.
        echo Automatic install failed. Launching manual installer...
        echo IMPORTANT: Check "Add Python to PATH" at the bottom!
        echo.
        "%PYTHON_INSTALLER%"
    )

    del "%PYTHON_INSTALLER%" 2>nul

    REM Refresh PATH for this session
    for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "USER_PATH=%%B"
    for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "SYS_PATH=%%B"
    set "PATH=%USER_PATH%;%SYS_PATH%"

    where python >nul 2>&1
    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Python installed but not found in PATH.
        echo Please close this window, open a NEW command prompt, and run build.bat again.
        pause
        exit /b 1
    )

    echo Python installed successfully!
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo   Found: %%i
)
echo.

REM -- Step 2: Check/upgrade pip --
echo [2/4] Checking pip...
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pip...
    python -m ensurepip --upgrade
)
python -m pip install --upgrade pip --quiet 2>nul
echo   pip OK
echo.

REM -- Step 3: Install dependencies --
echo [3/4] Installing dependencies...
echo   - requests
python -m pip install requests --quiet
echo   - pillow
python -m pip install pillow --quiet
echo   - openpyxl
python -m pip install openpyxl --quiet
echo   - pyinstaller
python -m pip install pyinstaller --quiet
echo   All dependencies installed.
echo.

REM -- Step 4: Build executable --
echo [4/4] Building executable...
echo   This may take 30-60 seconds...
echo.

if exist "build" rmdir /s /q build 2>nul
if exist "CCTVIPToolkit.spec" del "CCTVIPToolkit.spec" 2>nul

REM Clean dist but preserve user data files
if exist "dist\CCTVIPToolkit.exe" del "dist\CCTVIPToolkit.exe" 2>nul

REM Check for app icon
set "ICON_ARGS="
if not exist "app.ico" goto :no_icon
echo   Icon: app.ico found
set ICON_ARGS=--icon=app.ico --add-data "app.ico;."
goto :icon_done
:no_icon
echo   Icon: app.ico not found — building without icon
:icon_done

python -m PyInstaller --onefile --windowed ^
    --name "CCTVIPToolkit" ^
    --uac-admin ^
    %ICON_ARGS% ^
    --clean ^
    axis_toolkit_v3.py

echo.
echo ========================================
if exist "dist\CCTVIPToolkit.exe" (
    echo BUILD SUCCESSFUL!
    echo.
    REM Clean up build artifacts
    if exist "build" rmdir /s /q build 2>nul
    if exist "CCTVIPToolkit.spec" del "CCTVIPToolkit.spec" 2>nul
    REM Remove anything in dist that isn't the exe
    for %%F in ("dist\*.*") do (
        if /I not "%%~xF"==".exe" del "%%F" 2>nul
    )
    for /d %%D in ("dist\*") do rmdir /s /q "%%D" 2>nul
    for %%A in ("dist\CCTVIPToolkit.exe") do (
        set "SIZE=%%~zA"
        set /a "SIZE_MB=!SIZE! / 1048576"
        echo   File: dist\CCTVIPToolkit.exe [!SIZE_MB! MB]
    )
    echo.
    echo To distribute, just share:
    echo   dist\CCTVIPToolkit.exe
    echo.
    echo No Python or dependencies needed to RUN the exe.
    echo ========================================
) else (
    echo BUILD FAILED
    echo.
    echo Common fixes:
    echo   - Make sure axis_toolkit_v3.py is in this folder
    echo   - Make sure app.ico is in this folder
    echo   - Try: python -m PyInstaller --onefile axis_toolkit_v3.py
    echo ========================================
)
echo.
pause

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
    cctv_toolkit.py

echo.
echo ========================================
if exist "dist\CCTVIPToolkit.exe" (
    echo PYINSTALLER BUILD SUCCESSFUL!
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

    REM ============================================================
    REM Step 5: Build Inno Setup installer
    REM ============================================================
    echo.
    echo [5/5] Building installer...

    REM Pull APP_VERSION from cctv_toolkit.py so installer + .exe stay in sync
    call :read_version
    if "!APP_VER!"=="" (
        echo   WARN: could not read APP_VERSION from cctv_toolkit.py — defaulting to 0.0.0
        set "APP_VER=0.0.0"
    )
    echo   Version: !APP_VER!

    REM Pre-compute Inno Setup candidate paths OUTSIDE this block —
    REM literal "(x86)" in %ProgramFiles(x86)% breaks the cmd parser's paren matching
    REM when expanded inside an if-block.
    call :find_iscc
    if "!ISCC!"=="" (
        echo   Inno Setup 6 not installed — installing via winget...
        winget install --id JRSoftware.InnoSetup --accept-package-agreements --accept-source-agreements --silent
        call :find_iscc
    )

    if "!ISCC!"=="" (
        echo   ERROR: ISCC.exe not found after install attempt — installer skipped.
        echo   Install Inno Setup 6 manually from https://jrsoftware.org/isinfo.php
    ) else (
        echo   Compiler: !ISCC!
        if exist "Output" rmdir /s /q "Output" 2>nul
        "!ISCC!" /Q /DMyAppVersion=!APP_VER! installer.iss
        if exist "Output\CCTVIPToolkit-Setup-v!APP_VER!.exe" (
            move /Y "Output\CCTVIPToolkit-Setup-v!APP_VER!.exe" "dist\" >nul
            rmdir /s /q "Output" 2>nul
            for %%A in ("dist\CCTVIPToolkit-Setup-v!APP_VER!.exe") do (
                set "ISIZE=%%~zA"
                set /a "ISIZE_MB=!ISIZE! / 1048576"
                echo   File: dist\CCTVIPToolkit-Setup-v!APP_VER!.exe [!ISIZE_MB! MB]
            )
        ) else (
            echo   ERROR: Inno Setup compile failed — see output above.
        )
    )

    echo.
    echo ========================================
    echo BUILD COMPLETE
    echo.
    echo Distribute either or both:
    echo   dist\CCTVIPToolkit.exe                       ^(bare exe, runs from anywhere^)
    echo   dist\CCTVIPToolkit-Setup-v!APP_VER!.exe   ^(installer w/ Start Menu + uninstall^)
    echo ========================================
) else (
    echo BUILD FAILED
    echo.
    echo Common fixes:
    echo   - Make sure cctv_toolkit.py is in this folder
    echo   - Make sure app.ico is in this folder
    echo   - Try: python -m PyInstaller --onefile cctv_toolkit.py
    echo ========================================
)
echo.
pause
goto :eof

REM ============================================================
REM :find_iscc — sets ISCC to the first existing Inno Setup compiler path.
REM Defined as a callable label so the literal "(x86)" in the path
REM doesn't poison the cmd parenthesis parser inside if-blocks.
REM ============================================================
:find_iscc
set "ISCC="
set "_P1=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
set "_P2=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
set "_P3=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%_P1%" set "ISCC=%_P1%"
if "%ISCC%"=="" if exist "%_P2%" set "ISCC=%_P2%"
if "%ISCC%"=="" if exist "%_P3%" set "ISCC=%_P3%"
goto :eof

REM ============================================================
REM :read_version — sets APP_VER from APP_VERSION = "X.Y.Z" line in cctv_toolkit.py
REM Same isolation reason: complex quoting outside if-block.
REM ============================================================
:read_version
set "APP_VER="
for /f "tokens=2 delims==" %%V in ('findstr /B /C:"APP_VERSION" cctv_toolkit.py') do (
    set "RAW=%%V"
    REM Strip leading space and surrounding quotes
    set "RAW=!RAW: =!"
    set "RAW=!RAW:"=!"
    set "RAW=!RAW:'=!"
    set "APP_VER=!RAW!"
    goto :read_version_done
)
:read_version_done
goto :eof

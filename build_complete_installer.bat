@echo off
echo ========================================================================
echo  Building Complete Standalone Installer for Tray Serial Monitor Service
echo ========================================================================
echo.

REM Set error handling
setlocal enabledelayedexpansion

REM Step 1: Build standalone executables
echo [1/3] Building standalone executables with PyInstaller...
echo.
python build_executables.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Failed to build standalone executables!
    echo Make sure you have PyInstaller installed: pip install pyinstaller
    pause
    exit /b 1
)

echo.
echo ========================================================================

REM Step 2: Check if exe directory exists
if not exist "exe" (
    echo ERROR: exe directory not found! Executable build may have failed.
    pause
    exit /b 1
)

REM Check if Inno Setup is installed
echo [2/3] Checking Inno Setup installation...
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "INNO_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    goto :inno_found
)
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "INNO_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"
    goto :inno_found
)

echo ERROR: Inno Setup 6 not found!
echo Please install Inno Setup 6 from: https://jrsoftware.org/isinfo.php
echo.
echo Expected locations:
echo   C:\Program Files (x86)\Inno Setup 6\ISCC.exe
echo   C:\Program Files\Inno Setup 6\ISCC.exe
pause
exit /b 1

:inno_found
echo Found Inno Setup at: !INNO_PATH!

REM Step 3: Build the installer
echo.
echo [3/3] Building Windows installer...
echo.

REM Create output directory if it doesn't exist
if not exist "installer_output" mkdir installer_output

REM Build the installer
echo Compiling installer with Inno Setup...
"!INNO_PATH!" "service_installer.iss"

REM Check if installer was created successfully
if exist "installer_output\TraySerialMonitor_Standalone_Setup.exe" (
    echo.
    echo ========================================================================
    echo SUCCESS! Complete standalone installer created successfully!
    echo ========================================================================
    echo.
    echo Installer Location: installer_output\TraySerialMonitor_Standalone_Setup.exe
    echo.
    echo Executable Files Built:
    echo   - TrayHardwareMonitorService.exe  (Windows Service)
    echo   - TraySerialMonitorClient.exe     (GUI Client)
    echo   - InstallService.exe              (Service Installer)
    echo   - UninstallService.exe            (Service Uninstaller)
    echo.
    echo Ready for Distribution!
    echo The installer can be distributed to any Windows 10/11 system without
    echo requiring Python or any additional dependencies.
    echo.
    echo.
    echo Build process completed.
    pause
    exit /b 0
)

echo.
echo ========================================================================
echo ERROR: Failed to build installer!
echo ========================================================================
echo Check the Inno Setup output above for details.
echo.
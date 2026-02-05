@echo off
REM CopyAsInsert Auto-Start Setup Script
REM This script enables or disables auto-start functionality for CopyAsInsert

setlocal enabledelayedexpansion

echo.
echo ============================================
echo CopyAsInsert Auto-Start Setup
echo ============================================
echo.

REM Get the startup folder path
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

echo.
echo Choose an option:
echo 1. Enable auto-start (recommended)
echo 2. Disable auto-start
echo 3. Exit
echo.
choice /C 123 /M "Enter your choice (1-3): " /N

if errorlevel 3 goto exit
if errorlevel 2 goto disable
if errorlevel 1 goto enable

:enable
echo.
echo Locating CopyAsInsert.exe...

REM Try to find CopyAsInsert.exe in common locations
if exist "%CD%\CopyAsInsert.exe" (
    set "EXE_PATH=%CD%\CopyAsInsert.exe"
) else if exist "%ProgramFiles%\CopyAsInsert\CopyAsInsert.exe" (
    set "EXE_PATH=%ProgramFiles%\CopyAsInsert\CopyAsInsert.exe"
) else if exist "%LOCALAPPDATA%\CopyAsInsert\CopyAsInsert.exe" (
    set "EXE_PATH=%LOCALAPPDATA%\CopyAsInsert\CopyAsInsert.exe"
) else (
    echo.
    echo ERROR: CopyAsInsert.exe not found in:
    echo - Current directory
    echo - Program Files
    echo - AppData\Local
    echo.
    echo Please ensure CopyAsInsert.exe is installed and accessible.
    pause
    goto exit
)

echo Found: !EXE_PATH!
echo.

REM Create shortcut using PowerShell (more reliable than VBScript on modern Windows)
powershell -NoProfile -Command ^
    "$WshShell = New-Object -ComObject WScript.Shell; ^
    $Shortcut = $WshShell.CreateShortcut('%STARTUP_FOLDER%\CopyAsInsert.lnk'); ^
    $Shortcut.TargetPath = '!EXE_PATH!'; ^
    $Shortcut.WorkingDirectory = (Split-Path '!EXE_PATH!'); ^
    $Shortcut.WindowStyle = 7; ^
    $Shortcut.Save()" >nul 2>&1

if errorlevel 1 (
    echo.
    echo ERROR: Failed to create startup shortcut.
    echo Please check your permissions and try again.
    pause
    goto exit
)

echo.
echo SUCCESS: CopyAsInsert will now auto-start when Windows boots!
echo.
echo The shortcut has been created at:
echo %STARTUP_FOLDER%\CopyAsInsert.lnk
echo.
pause
goto exit

:disable
echo.
echo Removing auto-start shortcut...

if exist "%STARTUP_FOLDER%\CopyAsInsert.lnk" (
    del "%STARTUP_FOLDER%\CopyAsInsert.lnk" /Q
    if errorlevel 1 (
        echo ERROR: Failed to remove startup shortcut.
        echo.
        pause
        goto exit
    )
    echo SUCCESS: Auto-start has been disabled.
) else (
    echo INFO: Auto-start shortcut not found - nothing to remove.
)

echo.
pause
goto exit

:exit
echo.
endlocal

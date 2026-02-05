@echo off
REM CopyAsInsert Update Checker and Installer
REM This script checks if CopyAsInsert is installed and updates it if a new version is available

setlocal enabledelayedexpansion

cls
echo.
echo ============================================
echo CopyAsInsert Update Checker
echo ============================================
echo.

REM Check if CopyAsInsert.exe was provided as argument
if "%~1"=="" (
    echo ERROR: Please provide the path to the new CopyAsInsert.exe
    echo.
    echo Usage: UpdateCopyAsInsert.bat ^<path-to-new-exe^>
    echo.
    echo Example: UpdateCopyAsInsert.bat "C:\Downloads\CopyAsInsert.exe"
    echo.
    pause
    exit /b 1
)

set "NEW_EXE=%~1"
set "NEW_EXE_FULL=%CD%\%NEW_EXE%"

REM Resolve relative paths to absolute
if not "%NEW_EXE:~0,2%"=="\\" (
    if not "%NEW_EXE:~1,1%"==":" (
        set "NEW_EXE=%NEW_EXE_FULL%"
    )
)

REM Verify the new executable exists
if not exist "%NEW_EXE%" (
    echo ERROR: File not found: %NEW_EXE%
    echo.
    pause
    exit /b 1
)

echo New executable found: %NEW_EXE%
echo.

REM Try to find installed CopyAsInsert.exe
echo Searching for installed CopyAsInsert.exe...

set "FOUND_INSTALLATION="

if exist "%ProgramFiles%\CopyAsInsert\CopyAsInsert.exe" (
    set "FOUND_INSTALLATION=%ProgramFiles%\CopyAsInsert\CopyAsInsert.exe"
    echo [Program Files] Found at: !FOUND_INSTALLATION!
) else if exist "%LOCALAPPDATA%\CopyAsInsert\CopyAsInsert.exe" (
    set "FOUND_INSTALLATION=%LOCALAPPDATA%\CopyAsInsert\CopyAsInsert.exe"
    echo [AppData] Found at: !FOUND_INSTALLATION!
) else if exist "%CD%\CopyAsInsert.exe" (
    set "FOUND_INSTALLATION=%CD%\CopyAsInsert.exe"
    echo [Current Directory] Found at: !FOUND_INSTALLATION!
) else (
    echo.
    echo WARNING: No installed CopyAsInsert.exe found in standard locations.
    echo.
    echo Standard locations checked:
    echo - %ProgramFiles%\CopyAsInsert\
    echo - %LOCALAPPDATA%\CopyAsInsert\
    echo - Current directory
    echo.
    echo You can:
    echo 1. Proceed to use the new executable in its current location
    echo 2. Manually copy it to your preferred location
    echo.
    choice /C YN /M "Do you want to continue? (Y/N): " /N
    if errorlevel 2 goto exit
    if errorlevel 1 goto install_new
)

if not "!FOUND_INSTALLATION!"=="" (
    echo.
    echo Existing installation found at: !FOUND_INSTALLATION!
    echo.
    
    REM Extract version info if possible (check file properties)
    echo Preparing to update...
    echo.
    
    choice /C YN /M "Update existing installation? (Y/N): " /N
    if errorlevel 2 goto exit
    if errorlevel 1 goto perform_update
)

goto install_new

:perform_update
echo.
echo Stopping CopyAsInsert if running...

REM Kill the running process
taskkill /IM CopyAsInsert.exe /F /T >nul 2>&1
set "KILL_RESULT=!ERRORLEVEL!"

if !KILL_RESULT! equ 0 (
    echo Process stopped successfully
    timeout /t 2 /nobreak >nul
) else if !KILL_RESULT! equ 128 (
    echo Process was not running (OK)
) else (
    echo Warning: Could not stop process (continuing anyway)
)

echo.
echo Creating backup of current installation...

REM Create backup
set "BACKUP_PATH=!FOUND_INSTALLATION!.backup"
copy "!FOUND_INSTALLATION!" "!BACKUP_PATH!" /Y >nul 2>&1

if errorlevel 1 (
    echo Warning: Could not create backup
) else (
    echo Backup created at: !BACKUP_PATH!
)

echo.
echo Replacing executable...

REM Copy the new executable
copy "%NEW_EXE%" "!FOUND_INSTALLATION!" /Y >nul 2>&1

if errorlevel 1 (
    echo ERROR: Failed to copy new executable
    echo.
    if exist "!BACKUP_PATH!" (
        echo Restoring backup...
        copy "!BACKUP_PATH!" "!FOUND_INSTALLATION!" /Y >nul 2>&1
        echo Backup restored.
    )
    pause
    exit /b 1
)

echo.
echo SUCCESS: CopyAsInsert has been updated!
echo Location: !FOUND_INSTALLATION!
echo.
echo To restart the application:
echo 1. Open the updated CopyAsInsert.exe
echo 2. Or wait for auto-start if you have it enabled
echo.
pause
goto exit

:install_new
echo.
echo Installation Mode
echo ==================
echo.
echo CopyAsInsert.exe will be available at:
echo %NEW_EXE%
echo.
echo You can:
echo 1. Run it directly from here
echo 2. Create a shortcut on your Desktop
echo 3. Move it to Program Files or another location
echo.
echo Recommended: Run SetupAutoStart.bat after copying to your final location.
echo.
pause
goto exit

:exit
echo.
endlocal

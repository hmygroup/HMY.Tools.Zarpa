@echo off
REM Script to apply deferred update by replacing exe with staged version
REM Usage: ApplyUpdate.bat

setlocal enabledelayedexpansion

REM Get the directory where this script is running
set SCRIPT_DIR=%~dp0

REM Define file paths - work in the same directory as the script
set EXE_PATH=%SCRIPT_DIR%CopyAsInsert.exe
set NEW_EXE_PATH=%SCRIPT_DIR%CopyAsInsert.exe.new
set BACKUP_PATH=%SCRIPT_DIR%CopyAsInsert.exe.backup

REM Check if .new file exists
if not exist "!NEW_EXE_PATH!" (
    exit /b 0
)

REM Wait for processes to release file handles
timeout /t 2 /nobreak

REM Try to backup current exe
if exist "!EXE_PATH!" (
    if exist "!BACKUP_PATH!" (
        del "!BACKUP_PATH!" 2>nul
    )
    move /y "!EXE_PATH!" "!BACKUP_PATH!" >nul 2>&1
)

REM Move new exe to current location
if exist "!NEW_EXE_PATH!" (
    move /y "!NEW_EXE_PATH!" "!EXE_PATH!" >nul 2>&1
    if errorlevel 1 (
        REM If move failed, try to restore backup
        if exist "!BACKUP_PATH!" (
            move /y "!BACKUP_PATH!" "!EXE_PATH!" >nul 2>&1
        )
        exit /b 1
    )
)

exit /b 0

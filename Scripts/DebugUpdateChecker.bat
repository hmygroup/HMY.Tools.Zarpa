@echo off
REM CopyAsInsert - Update Checker Debug Tool
REM This script displays the update checker logs for troubleshooting

setlocal enabledelayedexpansion

set "LOG_FILE=%APPDATA%\CopyAsInsert\UpdateChecker.log"

echo.
echo ============================================
echo CopyAsInsert Update Checker - Debug Logs
echo ============================================
echo.

if not exist "%LOG_FILE%" (
    echo ERROR: Log file not found at:
    echo %LOG_FILE%
    echo.
    echo This means the app has not run yet or no update checks have been performed.
    echo.
    echo Run CopyAsInsert.exe first, then click "Check for Update" or wait for startup check.
    echo.
    pause
    exit /b 1
)

echo Log file: %LOG_FILE%
echo.
echo --- LATEST UPDATE CHECK LOG ---
echo.

REM Show last 50 lines
powershell -NoProfile -Command "Get-Content '%LOG_FILE%' -Tail 50 | Out-Host"

echo.
echo --- TROUBLESHOOTING ---
echo.
echo If you see errors:
echo.
echo 1. "tag_name property not found"
echo    - GitHub release might not be properly formatted
echo    - Check: https://github.com/hmygroup/HMY.Tools.Zarpa/releases
echo.
echo 2. "HTTP 404"
echo    - No releases published yet (expected if this is your first check)
echo    - Push to main branch to trigger GitHub Actions
echo.
echo 3. Network errors / timeouts
echo    - Your firewall / network might be blocking GitHub
echo    - Try accessing GitHub in your browser
echo.
echo 4. JSON parse error
echo    - GitHub API might have changed or returned invalid content
echo.
echo For more details, check the full log:
echo %LOG_FILE%
echo.
echo (Opening full log in Notepad)
echo.
pause

REM Open full log in Notepad
start notepad "%LOG_FILE%"

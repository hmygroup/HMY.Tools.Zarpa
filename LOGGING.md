# CopyAsInsert Logging Guide

## Overview
The application now includes comprehensive logging using **Serilog** to track all errors, warnings, and important events.

## Log File Location

Logs are saved to:
```
C:\Users\[YourUsername]\AppData\Roaming\CopyAsInsert\logs\log.txt
```

### Windows Path Examples
- On Windows: `%APPDATA%\CopyAsInsert\logs\log.txt`
- In File Explorer: Copy and paste the path mentioned above

## Log File Features

- **Daily Rolling Logs**: New log file created each day (log-20260205.txt, log-20260206.txt, etc.)
- **Retention**: Last 30 days of logs are kept automatically
- **Timestamp Format**: `[YYYY-MM-DD HH:mm:ss.fff zzz]`
- **Log Levels**: DEBUG, INFO, WARNING, ERROR, FATAL

## What Gets Logged

### Application Startup/Shutdown
- Application start time
- Hotkey registration status
- Application termination

### Data Processing Events
- Clipboard data parsing (TSV/CSV format detection)
- Excel file processing (column count, row count)
- Column type inference results
- SQL generation success/failure

### Errors and Exceptions
- All exceptions with full stack traces
- Error messages from forms and dialogs
- File access errors
- Parser errors with details

## Log Levels Explained

| Level | Use Case | Example |
|-------|----------|---------|
| **DEBUG** | Detailed diagnostic info | "ProcessClipboard started", "Column types inferred" |
| **INFO** | Important business events | "Clipboard parsed successfully: 5 columns, 100 rows", "SQL generated successfully" |
| **WARNING** | Warning conditions | Potential issues that don't stop operation |
| **ERROR** | Error events with exception | "ParseExcelFile failed: File not found" |
| **FATAL** | Critical failures | Unrecoverable errors |

## Example Log Output

```
[2026-02-05 14:32:18.123 +02:00] [INF] === CopyAsInsert Application Started ===
[2026-02-05 14:32:18.124 +02:00] [INF] Log directory: C:\Users\carlos.marin\AppData\Roaming\CopyAsInsert\logs
[2026-02-05 14:32:18.234 +02:00] [INF] Hotkey Alt+Shift+I registered successfully
[2026-02-05 14:35:42.567 +02:00] [DBG] ProcessClipboard started
[2026-02-05 14:35:42.568 +02:00] [INF] Clipboard contains tabular data, showing header check form
[2026-02-05 14:35:45.123 +02:00] [INF] Header check result: hasHeaders=True
[2026-02-05 14:35:45.125 +02:00] [INF] Clipboard parsed successfully: 3 columns, 50 rows
[2026-02-05 14:35:45.126 +02:00] [DBG] Column types inferred
[2026-02-05 14:35:48.456 +02:00] [INF] Config form accepted: TableName=Products, Schema=dbo
[2026-02-05 14:35:48.457 +02:00] [INF] SQL generated successfully: INSERT into Products with 50 rows
[2026-02-05 14:40:00.789 +02:00] [INF] === CopyAsInsert Application Ended ===
```

## Troubleshooting with Logs

### Problem: "Failed to parse clipboard data"
- Check logs for: `ParseClipboardText failed`
- Look for specific error message and exception details
- Verify clipboard format (TSV/CSV)

### Problem: "Excel file not processing correctly"
- Check logs for: `ParseExcelFile completed`
- Verify row count matches your expectations
- Look for warnings about empty rows being skipped

### Problem: "SQL generation errors"
- Check logs for: `SQL generation failed`
- Exception details will show what went wrong
- Look at column type inference results before the error

### Problem: "Hotkey not working"
- Check logs for: `Hotkey Alt+Shift+I registered successfully`
- If missing, look for: `Failed to register hotkey` with error details
- Restart the application

## Tips for Debugging

1. **Open the log file**:
   - Press `Win + R`
   - Type: `%APPDATA%\CopyAsInsert\logs`
   - Press Enter
   - Open `log.txt` with any text editor

2. **Monitor logs in real-time**:
   - Use `tail -f` in PowerShell/terminal to watch logs as they're written
   - Or use VS Code to open the log file and refresh it

3. **Search for errors**:
   - Search for `[ERR]` or `[FTL]` in the log file
   - Look for timestamps matching when the issue occurred

4. **Share logs for support**:
   - Provide the entire log for the session when reporting issues
   - Include the exact error message from the log

## Code Integration

The logging system is integrated throughout the application:

- **Logger.cs**: Central logging service (Services folder)
- **MainForm.cs**: Logs UI events, hotkey registration, clipboard processing
- **TableDataParser.cs**: Logs data parsing operations
- **Services**: All service classes log important operations

## Log File Size Management

- Log files automatically rotate daily
- Each day gets a new file: `log-20260205.txt`, `log-20260206.txt`, etc.
- Old logs (older than 30 days) are automatically deleted
- No manual cleanup needed

## Disabling or Changing Logs

To change log settings, edit `Logger.cs`:
- Change `MinimumLevel.Debug()` to `MinimumLevel.Information()` to reduce verbosity
- Modify file path in `logPath` variable
- Adjust `retainedFileCountLimit` to keep more/fewer days of logs
- Change `rollingInterval` for different rotation periods (Hourly, Weekly, etc.)

---

**Last Updated**: February 5, 2026
**Logger Version**: Serilog 4.0.0

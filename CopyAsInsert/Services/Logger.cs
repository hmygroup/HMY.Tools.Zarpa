using System;
using Serilog;
using Serilog.Core;

namespace CopyAsInsert.Services;

/// <summary>
/// Centralized logging service for the application
/// </summary>
public static class Logger
{
    private static ILogger? _logger;
    private static readonly object _lock = new object();
    // Event for UI components to subscribe and receive log messages in real-time
    public static event Action<string, string>? MessageLogged;

    /// <summary>
    /// Initialize the logging system
    /// </summary>
    public static void Initialize()
    {
        lock (_lock)
        {
            if (_logger != null)
                return; // Already initialized

            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CopyAsInsert",
                "logs",
                "log.txt"
            );

            // Create directories if they don't exist
            Directory.CreateDirectory(Path.GetDirectoryName(logPath) ?? "");

            _logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    logPath,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz}] [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 30
                )
                .CreateLogger();

            LogInfo("=== CopyAsInsert Application Started ===");
            LogInfo($"Log directory: {Path.GetDirectoryName(logPath)}");
        }
    }

    /// <summary>
    /// Get the underlying Serilog logger instance
    /// </summary>
    public static ILogger GetLogger()
    {
        if (_logger == null)
            Initialize();

        return _logger ?? throw new InvalidOperationException("Logger failed to initialize");
    }

    /// <summary>
    /// Log info message
    /// </summary>
    public static void LogInfo(string message)
    {
        GetLogger().Information(message);
        try { MessageLogged?.Invoke("INFO", message); } catch { }
    }

    /// <summary>
    /// Log debug message
    /// </summary>
    public static void LogDebug(string message)
    {
        GetLogger().Debug(message);
        try { MessageLogged?.Invoke("DBG", message); } catch { }
    }

    /// <summary>
    /// Log warning message
    /// </summary>
    public static void LogWarning(string message)
    {
        GetLogger().Warning(message);
        try { MessageLogged?.Invoke("WRN", message); } catch { }
    }

    /// <summary>
    /// Log error with exception
    /// </summary>
    public static void LogError(string message, Exception ex)
    {
        GetLogger().Error(ex, message);
        try { MessageLogged?.Invoke("ERR", message + " - " + ex.ToString()); } catch { }
    }

    /// <summary>
    /// Log error without exception
    /// </summary>
    public static void LogError(string message)
    {
        GetLogger().Error(message);
        try { MessageLogged?.Invoke("ERR", message); } catch { }
    }

    /// <summary>
    /// Log fatal error
    /// </summary>
    public static void LogFatal(string message, Exception ex)
    {
        GetLogger().Fatal(ex, message);
        try { MessageLogged?.Invoke("FTL", message + " - " + ex.ToString()); } catch { }
    }

    /// <summary>
    /// Close and flush logger
    /// </summary>
    public static void CloseAndFlush()
    {
        try
        {
            if (_logger != null)
            {
                LogInfo("=== CopyAsInsert Application Ended ===");
                Log.CloseAndFlush();
            }
        }
        catch
        {
            // Silently fail if logger is already closed
        }
    }
}

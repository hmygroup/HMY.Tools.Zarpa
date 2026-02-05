using Microsoft.Win32;

namespace CopyAsInsert.Services;

/// <summary>
/// Manages Windows Registry entries for application startup
/// </summary>
public class StartupManager
{
    private static readonly string STARTUP_REGISTRY_PATH = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private static readonly string APPLICATION_NAME = "CopyAsInsert";

    /// <summary>
    /// Enable application to run on Windows startup
    /// </summary>
    public static bool EnableStartup()
    {
        try
        {
            string exePath = Application.ExecutablePath ?? "";
            
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(STARTUP_REGISTRY_PATH, true))
            {
                if (key != null)
                {
                    key.SetValue(APPLICATION_NAME, exePath);
                    Logger.LogInfo($"Startup enabled: {exePath}");
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to enable startup in Registry: {ex.Message}");
        }

        return false;
    }

    /// <summary>
    /// Disable application from running on Windows startup
    /// </summary>
    public static bool DisableStartup()
    {
        try
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(STARTUP_REGISTRY_PATH, true))
            {
                if (key != null)
                {
                    var value = key.GetValue(APPLICATION_NAME);
                    if (value != null)
                    {
                        key.DeleteValue(APPLICATION_NAME);
                        Logger.LogInfo("Startup disabled");
                    }
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to disable startup in Registry: {ex.Message}");
        }

        return false;
    }

    /// <summary>
    /// Check if application is set to run on startup
    /// </summary>
    public static bool IsStartupEnabled()
    {
        try
        {
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(STARTUP_REGISTRY_PATH))
            {
                if (key != null)
                {
                    return key.GetValue(APPLICATION_NAME) != null;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Failed to check startup status: {ex.Message}");
        }

        return false;
    }
}

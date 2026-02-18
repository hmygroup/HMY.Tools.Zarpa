using System.Text.Json;
using CopyAsInsert.Models;

namespace CopyAsInsert.Services;

/// <summary>
/// Manages application settings persistence to JSON file
/// </summary>
public class SettingsManager
{
    private static readonly string SettingsDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "CopyAsInsert"
    );

    private static readonly string SettingsPath = Path.Combine(SettingsDirectory, "settings.json");
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public class ApplicationSettings
    {
        public string DefaultSchema { get; set; } = "dbo";
        public bool AutoCreateHistoryTable { get; set; } = true;
        public bool TemporalTableByDefault { get; set; } = true;
        public bool RunOnStartup { get; set; } = false;
        public int HotKeyModifier { get; set; } = 0x0001 | 0x0004; // MOD_ALT | MOD_SHIFT
        public int HotKeyVirtualKey { get; set; } = 0x49; // 'I'
        public bool AutoAppendTemporalSuffix { get; set; } = false; // Control "_Temporal" suffix
        public bool ShowFormOnStartup { get; set; } = false; // Show main form on startup
    }

    /// <summary>
    /// Load settings from JSON file
    /// </summary>
    public static ApplicationSettings LoadSettings()
    {
        try
        {
            if (!Directory.Exists(SettingsDirectory))
            {
                Directory.CreateDirectory(SettingsDirectory);
            }

            if (File.Exists(SettingsPath))
            {
                string json = File.ReadAllText(SettingsPath);
                var settings = JsonSerializer.Deserialize<ApplicationSettings>(json, JsonOptions);
                if (settings != null)
                {
                    Logger.LogDebug($"Settings loaded from {SettingsPath}");
                    return settings;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Failed to load settings: {ex.Message}");
        }

        // Return defaults if file doesn't exist or parsing fails
        return new ApplicationSettings();
    }

    /// <summary>
    /// Save settings to JSON file
    /// </summary>
    public static void SaveSettings(ApplicationSettings settings)
    {
        try
        {
            if (!Directory.Exists(SettingsDirectory))
            {
                Directory.CreateDirectory(SettingsDirectory);
            }

            string json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(SettingsPath, json);
            Logger.LogDebug($"Settings saved to {SettingsPath}");
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to save settings: {ex.Message}");
        }
    }

    /// <summary>
    /// Get the settings file path
    /// </summary>
    public static string GetSettingsPath() => SettingsPath;
}

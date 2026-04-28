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
    private const int MaxRecentExcelValues = 15;

    public class ApplicationSettings
    {
        public string DefaultSchema { get; set; } = "dbo";
        public bool AutoCreateHistoryTable { get; set; } = true;
        public bool TemporalTableByDefault { get; set; } = false;
        public bool RunOnStartup { get; set; } = false;
        public int HotKeyModifier { get; set; } = 0x0001 | 0x0004; // MOD_ALT | MOD_SHIFT
        public int HotKeyVirtualKey { get; set; } = 0x49; // 'I'
        public bool AutoAppendTemporalSuffix { get; set; } = false; // Control "_Temporal" suffix
        public bool ShowFormOnStartup { get; set; } = false; // Show main form on startup
        
        // Excel Import settings
        public string ExcelImportServer { get; set; } = string.Empty; // Last used SQL Server
        public string ExcelImportDatabase { get; set; } = string.Empty; // Last used database
        public List<string> ExcelImportServerHistory { get; set; } = new();
        public List<string> ExcelImportDatabaseHistory { get; set; } = new();
        public int ExcelImportHotKeyModifier { get; set; } = 0x0001 | 0x0004; // MOD_ALT | MOD_SHIFT
        public int ExcelImportHotKeyVirtualKey { get; set; } = 0x45; // 'E'
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
                    return NormalizeExcelImportSettings(settings);
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

            settings = NormalizeExcelImportSettings(settings);
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
    /// Update last used Excel import server/database and keep recent values ordered by most recent.
    /// </summary>
    internal static void UpdateExcelImportRecentValues(ApplicationSettings settings, string? server, string? database)
    {
        settings.ExcelImportServer = server?.Trim() ?? string.Empty;
        settings.ExcelImportDatabase = database?.Trim() ?? string.Empty;
        settings.ExcelImportServerHistory = NormalizeRecentValues(settings.ExcelImportServerHistory, settings.ExcelImportServer);
        settings.ExcelImportDatabaseHistory = NormalizeRecentValues(settings.ExcelImportDatabaseHistory, settings.ExcelImportDatabase);
    }

    /// <summary>
    /// Get the settings file path
    /// </summary>
    public static string GetSettingsPath() => SettingsPath;

    private static ApplicationSettings NormalizeExcelImportSettings(ApplicationSettings settings)
    {
        settings.ExcelImportServer = settings.ExcelImportServer?.Trim() ?? string.Empty;
        settings.ExcelImportDatabase = settings.ExcelImportDatabase?.Trim() ?? string.Empty;
        settings.ExcelImportServerHistory = NormalizeRecentValues(settings.ExcelImportServerHistory, settings.ExcelImportServer);
        settings.ExcelImportDatabaseHistory = NormalizeRecentValues(settings.ExcelImportDatabaseHistory, settings.ExcelImportDatabase);

        if (string.IsNullOrWhiteSpace(settings.ExcelImportServer) && settings.ExcelImportServerHistory.Count > 0)
        {
            settings.ExcelImportServer = settings.ExcelImportServerHistory[0];
        }

        if (string.IsNullOrWhiteSpace(settings.ExcelImportDatabase) && settings.ExcelImportDatabaseHistory.Count > 0)
        {
            settings.ExcelImportDatabase = settings.ExcelImportDatabaseHistory[0];
        }

        return settings;
    }

    internal static List<string> NormalizeRecentValues(IEnumerable<string>? existingValues, string? preferredValue)
    {
        List<string> normalizedValues = new();

        AddRecentValue(normalizedValues, preferredValue);

        if (existingValues != null)
        {
            foreach (string? value in existingValues)
            {
                AddRecentValue(normalizedValues, value);
                if (normalizedValues.Count >= MaxRecentExcelValues)
                {
                    break;
                }
            }
        }

        return normalizedValues;
    }

    private static void AddRecentValue(List<string> target, string? value)
    {
        string normalizedValue = value?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalizedValue))
        {
            return;
        }

        if (target.Contains(normalizedValue, StringComparer.OrdinalIgnoreCase))
        {
            return;
        }

        target.Add(normalizedValue);
    }

}

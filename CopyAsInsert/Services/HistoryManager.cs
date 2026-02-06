using System.Text.Json;
using CopyAsInsert.Models;

namespace CopyAsInsert.Services;

/// <summary>
/// Manages conversion history persistence to JSON file
/// </summary>
public class HistoryManager
{
    private static readonly string HistoryDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "CopyAsInsert"
    );

    private static readonly string HistoryPath = Path.Combine(HistoryDirectory, "history.json");
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    /// Load conversion history from JSON file
    /// </summary>
    public static List<ConversionResult> LoadHistory()
    {
        try
        {
            if (!Directory.Exists(HistoryDirectory))
            {
                Directory.CreateDirectory(HistoryDirectory);
            }

            if (File.Exists(HistoryPath))
            {
                string json = File.ReadAllText(HistoryPath);
                var history = JsonSerializer.Deserialize<List<ConversionResult>>(json, JsonOptions);
                if (history != null)
                {
                    Logger.LogDebug($"History loaded from {HistoryPath}: {history.Count} items");
                    return history;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Failed to load history: {ex.Message}");
        }

        // Return empty list if file doesn't exist or parsing fails
        return new List<ConversionResult>();
    }

    /// <summary>
    /// Save conversion history to JSON file
    /// </summary>
    public static void SaveHistory(List<ConversionResult> history)
    {
        try
        {
            if (!Directory.Exists(HistoryDirectory))
            {
                Directory.CreateDirectory(HistoryDirectory);
            }

            string json = JsonSerializer.Serialize(history, JsonOptions);
            File.WriteAllText(HistoryPath, json);
            Logger.LogDebug($"History saved to {HistoryPath}: {history.Count} items");
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to save history: {ex.Message}");
        }
    }

    /// <summary>
    /// Get the history file path
    /// </summary>
    public static string GetHistoryPath() => HistoryPath;
}

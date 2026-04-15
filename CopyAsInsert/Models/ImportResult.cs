namespace CopyAsInsert.Models;

/// <summary>
/// Result of a SQL Server query import to Excel operation
/// </summary>
public class ImportResult
{
    public bool Success { get; set; }
    
    /// <summary>
    /// Number of rows imported
    /// </summary>
    public int RowCount { get; set; }
    
    /// <summary>
    /// Path to the generated Excel file
    /// </summary>
    public string? OutputPath { get; set; }
    
    /// <summary>
    /// Error message if Success is false
    /// </summary>
    public string? ErrorMessage { get; set; }
    
    /// <summary>
    /// Stack trace for detailed error debugging
    /// </summary>
    public string? ErrorStackTrace { get; set; }
    
    /// <summary>
    /// Timestamp of import
    /// </summary>
    public DateTime ImportTime { get; set; } = DateTime.Now;
    
    /// <summary>
    /// Server name used for the import
    /// </summary>
    public string? ServerName { get; set; }
    
    /// <summary>
    /// Database name used for the import
    /// </summary>
    public string? DatabaseName { get; set; }
    
    /// <summary>
    /// Summary for display (e.g., "50 rows from Orders")
    /// </summary>
    public string Summary => $"{RowCount} rows from {DatabaseName}";
}

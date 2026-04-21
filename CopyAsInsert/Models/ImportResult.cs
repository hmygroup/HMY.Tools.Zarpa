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
    /// Number of Excel queries created during the import
    /// </summary>
    public int QueryCount { get; set; } = 1;
    
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

    /// <summary>
    /// Excel window handle (if available) returned by the import operation
    /// </summary>
    public int? ExcelHwnd { get; set; }

    /// <summary>
    /// True if the import operation attempted to bring Excel to the foreground
    /// </summary>
    public bool ExcelBroughtToForeground { get; set; }
}

namespace CopyAsInsert.Models;

/// <summary>
/// Result of a SQL conversion operation
/// </summary>
public class ConversionResult
{
    public bool Success { get; set; }
    
    /// <summary>
    /// Generated SQL script (CREATE TABLE + INSERT INTO statements)
    /// </summary>
    public string GeneratedSql { get; set; } = string.Empty;
    
    /// <summary>
    /// Number of rows affected/inserted
    /// </summary>
    public int RowCount { get; set; }
    
    /// <summary>
    /// Table name used in generation
    /// </summary>
    public string TableName { get; set; } = string.Empty;
    
    /// <summary>
    /// Schema used (e.g., dbo)
    /// </summary>
    public string Schema { get; set; } = "dbo";
    
    /// <summary>
    /// Error message if Success is false
    /// </summary>
    public string? ErrorMessage { get; set; }
    
    /// <summary>
    /// Timestamp of conversion
    /// </summary>
    public DateTime ConversionTime { get; set; } = DateTime.Now;
    
    /// <summary>
    /// Summary for display (e.g., "Orders table, 15 rows")
    /// </summary>
    public string Summary => $"{TableName} ({RowCount} rows)";
}

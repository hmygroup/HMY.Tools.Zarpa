namespace CopyAsInsert.Models;

/// <summary>
/// Represents inferred type information for a single column
/// </summary>
public class ColumnTypeInfo
{
    public required string ColumnName { get; set; }
    
    /// <summary>
    /// SQL type: INT, FLOAT, VARCHAR, DATETIME2, BOOL
    /// </summary>
    public required string SqlType { get; set; }
    
    /// <summary>
    /// Confidence percentage (0-100) based on lenient matching (70%+ threshold)
    /// </summary>
    public int ConfidencePercent { get; set; }
    
    /// <summary>
    /// Sample value from the column for preview
    /// </summary>
    public string SampleValue { get; set; } = string.Empty;
    
    /// <summary>
    /// Whether this column is detected as primary key
    /// </summary>
    public bool IsPrimaryKey { get; set; }
    
    /// <summary>
    /// Whether this column allows NULL values
    /// </summary>
    public bool AllowNull { get; set; }
    
    /// <summary>
    /// Maximum length for VARCHAR columns
    /// </summary>
    public int? MaxLength { get; set; }
}

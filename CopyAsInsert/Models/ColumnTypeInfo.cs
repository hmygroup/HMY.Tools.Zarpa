namespace CopyAsInsert.Models;

/// <summary>
/// Represents inferred type information for a single column
/// </summary>
public class ColumnTypeInfo
{
    public required string ColumnName { get; set; }
    
    /// <summary>
    /// SQL type: INT, FLOAT, DATETIME2, NVARCHAR
    /// </summary>
    public required string SqlType { get; set; }
    
    /// <summary>
    /// Confidence percentage (0-100) based on type matching
    /// Represents the percentage of non-empty values matching the inferred type
    /// </summary>
    public int ConfidencePercent { get; set; }
    
    /// <summary>
    /// Precise confidence score as a decimal (0.0-1.0)
    /// Useful for sorting and detailed analysis
    /// </summary>
    public double ConfidenceScore { get; set; }
    
    /// <summary>
    /// Human-readable explanation of why this type was selected
    /// Examples: "85% of values are valid integers", "ISO date format detected"
    /// </summary>
    public string InferenceReason { get; set; } = string.Empty;
    
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
    /// Maximum length for NVARCHAR columns
    /// </summary>
    public int? MaxLength { get; set; }
    
    /// <summary>
    /// Supported type options for user override (INT, FLOAT, DATETIME2, NVARCHAR)
    /// </summary>
    public string[] SupportedTypes { get; set; } = new[] { "INT", "FLOAT", "DATETIME2", "NVARCHAR" };
}

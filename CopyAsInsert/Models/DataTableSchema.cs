namespace CopyAsInsert.Models;

/// <summary>
/// Represents the parsed table structure with inferred types
/// </summary>
public class DataTableSchema
{
    public string OriginalTableName { get; set; } = "ParsedTable";
    
    /// <summary>
    /// Column type information for all columns, in order
    /// </summary>
    public List<ColumnTypeInfo> Columns { get; set; } = new();
    
    /// <summary>
    /// Raw data rows (strings), corresponding to columns
    /// </summary>
    public List<string[]> DataRows { get; set; } = new();
    
    /// <summary>
    /// Source of data: "ClipboardTSV", "ClipboardCSV", or "ExcelFile"
    /// </summary>
    public string DataSource { get; set; } = "ClipboardTSV";
    
    /// <summary>
    /// Column index of the detected primary key (or -1 if none)
    /// </summary>
    public int PrimaryKeyColumnIndex { get; set; } = -1;
    
    /// <summary>
    /// Allow users to override inferred types before SQL generation
    /// </summary>
    public bool AllowUserOverride { get; set; } = true;
    
    /// <summary>
    /// Total number of data rows (excluding header)
    /// </summary>
    public int RowCount => DataRows.Count;
}

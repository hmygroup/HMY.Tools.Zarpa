using CopyAsInsert.Models;
using System.Text;

namespace CopyAsInsert.Services;

/// <summary>
/// Generates SQL Server temporal table CREATE TABLE and INSERT statements
/// </summary>
public class SqlServerGenerator
{
    /// <summary>
    /// Generate CREATE TABLE and INSERT statements for temporal table
    /// </summary>
    public static ConversionResult GenerateSql(DataTableSchema schema, string tableName, string schema_name, bool isTemporalTable, bool isTemporaryTable = false, bool autoAppendTemporalSuffix = false)
    {
        var result = new ConversionResult
        {
            TableName = tableName,
            Schema = schema_name,
            RowCount = schema.RowCount
        };

        try
        {
            ValidateTableName(tableName);
            
            var sql = new StringBuilder();

            // Generate CREATE TABLE statement
            sql.Append(GenerateCreateTableStatement(schema, tableName, schema_name, isTemporalTable, isTemporaryTable, autoAppendTemporalSuffix));
            sql.AppendLine();
            sql.AppendLine();

            // Generate INSERT statements (one per row)
            sql.Append(GenerateInsertStatements(schema, tableName, schema_name, isTemporaryTable, autoAppendTemporalSuffix));

            result.GeneratedSql = sql.ToString();
            result.Success = true;
            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
            return result;
        }
    }

    /// <summary>
    /// Generate CREATE TABLE statement with temporal table features
    /// </summary>
    private static string GenerateCreateTableStatement(DataTableSchema schema, string tableName, string schema_name, bool isTemporalTable, bool isTemporaryTable = false, bool autoAppendTemporalSuffix = false)
    {
        var sb = new StringBuilder();

        string fullTableName;
        string historyTableName;

        // Determine table name suffix based on settings
        string suffix = (isTemporalTable && autoAppendTemporalSuffix) ? "_Temporal" : "";

        if (isTemporaryTable)
        {
            // Session-scoped temporary table uses # prefix
            fullTableName = $"[#{tableName}{suffix}]";
            historyTableName = $"[#{tableName}_History]";
        }
        else
        {
            // Regular table uses schema
            fullTableName = $"[{schema_name}].[{tableName}{suffix}]";
            historyTableName = $"[{schema_name}].[{tableName}_History]";
        }

        sb.AppendLine($"CREATE TABLE {fullTableName}");
        sb.AppendLine("(");

        // Add data columns
        for (int i = 0; i < schema.Columns.Count; i++)
        {
            var col = schema.Columns[i];
            var colDef = GenerateColumnDefinition(col);
            
            if (i < schema.Columns.Count - 1)
                sb.AppendLine($"    {colDef},");
            else
                sb.AppendLine($"    {colDef},");
        }

        if (isTemporalTable)
        {
            // Add temporal columns
            sb.AppendLine("    SysStartTime DATETIME2 GENERATED ALWAYS AS ROW START NOT NULL,");
            sb.AppendLine("    SysEndTime DATETIME2 GENERATED ALWAYS AS ROW END NOT NULL,");
            sb.AppendLine("    PERIOD FOR SYSTEM_TIME (SysStartTime, SysEndTime)");
        }

        sb.AppendLine(")");

        if (isTemporalTable)
        {
            sb.AppendLine($"WITH (SYSTEM_VERSIONING = ON (HISTORY_TABLE = {historyTableName}, DATA_CONSISTENCY_CHECK = ON));");
        }
        else
        {
            sb.AppendLine(";");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Generate column definition SQL - supports INT, FLOAT, DATETIME2, BIT, and NVARCHAR
    /// </summary>
    private static string GenerateColumnDefinition(ColumnTypeInfo col)
    {
        var def = new StringBuilder();
        def.Append($"[{col.ColumnName}] ");

        // Add SQL type based on inferred type
        switch (col.SqlType)
        {
            case "INT":
                def.Append("INT");
                break;
            
            case "FLOAT":
                // Use decimal(18,4) for better precision and compatibility
                def.Append("DECIMAL(18, 4)");
                break;
            
            case "DATETIME2":
                def.Append("DATETIME2(7)");
                break;
            
            case "BIT":
                def.Append("BIT");
                break;
            
            case "NVARCHAR":
            default:
                int length = (col.MaxLength.HasValue && col.MaxLength > 0) ? col.MaxLength.Value : 255;
                // Cap length at reasonable limit, use MAX for very long values
                if (length > 4000)
                    def.Append("NVARCHAR(MAX)");
                else
                    def.Append($"NVARCHAR({length})");
                break;
        }

        // Always allow NULL for columns
        def.Append(" NULL");

        return def.ToString();
    }

    /// <summary>
    /// Generate INSERT INTO statements (batched in groups of 1000)
    /// Includes ALL columns - no exclusions
    /// Must match the table name used in CREATE TABLE
    /// </summary>
    private static string GenerateInsertStatements(DataTableSchema schema, string tableName, string schema_name, bool isTemporaryTable = false, bool autoAppendTemporalSuffix = false)
    {
        var sb = new StringBuilder();

        // Table name must match CREATE TABLE exactly
        // Use the same suffix logic as CREATE TABLE
        string suffix = autoAppendTemporalSuffix ? "_Temporal" : "";
        
        string fullTableName;
        if (isTemporaryTable)
        {
            fullTableName = $"[#{tableName}{suffix}]";
        }
        else
        {
            fullTableName = $"[{schema_name}].[{tableName}{suffix}]";
        }
        
        // Build column list - ALL columns, no skipping
        var insertColumns = new List<string>();
        for (int i = 0; i < schema.Columns.Count; i++)
        {
            insertColumns.Add($"[{schema.Columns[i].ColumnName}]");
        }

        var columnList = string.Join(", ", insertColumns);
        Logger.LogDebug($"Insert columns ({insertColumns.Count}): {columnList}");

        const int BATCH_SIZE = 1000;
        int totalRows = schema.DataRows.Count;
        int batchCount = (totalRows + BATCH_SIZE - 1) / BATCH_SIZE;

        for (int batchIndex = 0; batchIndex < batchCount; batchIndex++)
        {
            int batchStart = batchIndex * BATCH_SIZE;
            int batchEnd = Math.Min(batchStart + BATCH_SIZE, totalRows);

            // Start new INSERT for this batch
            sb.AppendLine($"INSERT INTO {fullTableName} ({columnList})");
            sb.AppendLine("VALUES");

            // Generate VALUES for this batch
            var batchValues = new List<string>();
            for (int rowIndex = batchStart; rowIndex < batchEnd; rowIndex++)
            {
                var row = schema.DataRows[rowIndex];
                var values = new List<string>();

                // Add ALL column values in order
                for (int colIndex = 0; colIndex < schema.Columns.Count; colIndex++)
                {
                    var col = schema.Columns[colIndex];
                    var value = row[colIndex];
                    values.Add(FormatSqlValue(value, col.SqlType));
                }

                var valuesStr = string.Join(", ", values);
                batchValues.Add($"    ({valuesStr})");
            }

            sb.AppendLine(string.Join("," + Environment.NewLine, batchValues));
            sb.AppendLine(";");

            // Add separator between batches for clarity
            if (batchIndex < batchCount - 1)
                sb.AppendLine();
        }

        return sb.ToString();
    }

    /// <summary>
    /// Format a value for SQL based on its type - supports INT, FLOAT, DATETIME2, BIT, NVARCHAR
    /// Handles European decimal format (comma separator) for number types
    /// </summary>
    private static string FormatSqlValue(string value, string sqlType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "NULL";

        // If the value is exactly "NULL" (case-insensitive), treat it as the NULL keyword
        if (value.Equals("NULL", StringComparison.OrdinalIgnoreCase))
            return "NULL";

        return sqlType switch
        {
            "INT" => FormatIntValue(value),
            
            "FLOAT" => FormatFloatValue(value),
            
            "DATETIME2" => FormatDateTimeValue(value),
            
            "BIT" => FormatBooleanValue(value),
            
            _ => $"'{EscapeSqlString(value)}'"  // NVARCHAR or default
        };
    }

    /// <summary>
    /// Format integer value, handling European number format
    /// </summary>
    private static string FormatIntValue(string value)
    {
        // Remove thousand separators if present
        string normalized = value.Replace(".", "").Replace(",", "");
        
        if (int.TryParse(normalized, out var intVal))
            return intVal.ToString();
        
        // Try as long for larger numbers
        if (long.TryParse(normalized, out var longVal))
            return longVal.ToString();
        
        return "NULL";
    }

    /// <summary>
    /// Format float value, handling European decimal format (comma separator)
    /// </summary>
    private static string FormatFloatValue(string value)
    {
        // Normalize European format (1,5) to standard format (1.5)
        string normalized = value.Replace(".", "").Replace(",", ".");
        
        if (decimal.TryParse(normalized, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var floatVal))
            return floatVal.ToString(System.Globalization.CultureInfo.InvariantCulture);
        
        return "NULL";
    }

    /// <summary>
    /// Format a value as a DATETIME2 literal or NULL
    /// </summary>
    private static string FormatDateTimeValue(string value)
    {
        if (DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out var dt) ||
            DateTime.TryParse(value, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out dt))
        {
            return $"'{dt:yyyy-MM-dd HH:mm:ss.fffffff}'";
        }
        return "NULL";
    }

    /// <summary>
    /// Format a value as a BIT literal (0 or 1) or NULL
    /// </summary>
    private static string FormatBooleanValue(string value)
    {
        var booleanTrue = new[] { "true", "yes", "1", "on", "t", "y" };
        var booleanFalse = new[] { "false", "no", "0", "off", "f", "n" };
        
        string lower = value.ToLowerInvariant().Trim();
        
        // Remove any decimal separators for comparison
        string cleaned = lower.Replace(".", "").Replace(",", "");
        
        if (booleanTrue.Contains(cleaned))
            return "1";
        if (booleanFalse.Contains(cleaned))
            return "0";
        
        return "NULL";
    }

    /// <summary>
    /// Escape single quotes in SQL string values
    /// </summary>
    private static string EscapeSqlString(string value)
    {
        return value.Replace("'", "''");
    }

    /// <summary>
    /// Validate table name against SQL Server naming rules
    /// </summary>
    private static void ValidateTableName(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            throw new ArgumentException("Table name cannot be empty");

        if (tableName.Length > 128)
            throw new ArgumentException("Table name cannot exceed 128 characters");

        // Check for valid characters (alphanumeric, underscore, @, #)
        if (!System.Text.RegularExpressions.Regex.IsMatch(tableName, @"^[a-zA-Z_@#][a-zA-Z0-9_@#]*$"))
            throw new ArgumentException("Table name contains invalid characters");

        // Check for reserved keywords
        var reserved = new[] { "SELECT", "INSERT", "UPDATE", "DELETE", "CREATE", "ALTER", "DROP", "TABLE", "DATABASE", "VIEW" };
        if (reserved.Contains(tableName.ToUpper()))
            throw new ArgumentException("Table name is a reserved SQL keyword");
    }
}

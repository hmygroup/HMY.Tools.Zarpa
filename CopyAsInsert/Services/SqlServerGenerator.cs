using CopyAsInsert.Models;
using System.Text;
using System.Globalization;

namespace CopyAsInsert.Services;

/// <summary>
/// Generates SQL Server temporal table CREATE TABLE and INSERT statements
/// </summary>
public class SqlServerGenerator
{
    /// <summary>
    /// Generate CREATE TABLE and INSERT statements for temporal table
    /// </summary>
    public static ConversionResult GenerateSql(DataTableSchema schema, string tableName, string schema_name, bool isTemporalTable, bool isTemporaryTable = false)
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
            sql.Append(GenerateCreateTableStatement(schema, tableName, schema_name, isTemporalTable, isTemporaryTable));
            sql.AppendLine();
            sql.AppendLine();

            // Generate INSERT statements (one per row)
            sql.Append(GenerateInsertStatements(schema, tableName, schema_name, isTemporaryTable));

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
    private static string GenerateCreateTableStatement(DataTableSchema schema, string tableName, string schema_name, bool isTemporalTable, bool isTemporaryTable = false)
    {
        var sb = new StringBuilder();

        string fullTableName;
        string historyTableName;

        if (isTemporaryTable)
        {
            // Session-scoped temporary table uses # prefix
            fullTableName = $"[#{tableName}_Temporal]";
            historyTableName = $"[#{tableName}_History]";
        }
        else
        {
            // Regular table uses schema
            fullTableName = $"[{schema_name}].[{tableName}_Temporal]";
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
    /// Generate column definition SQL
    /// Uses only basic types: INT, DECIMAL, MONEY, NVARCHAR
    /// </summary>
    private static string GenerateColumnDefinition(ColumnTypeInfo col)
    {
        var def = new StringBuilder();
        def.Append($"[{col.ColumnName}] ");

        // Add SQL type - only basic types
        if (col.SqlType == "INT")
        {
            if (col.IsPrimaryKey)
                def.Append("INT IDENTITY(1,1) PRIMARY KEY");
            else
                def.Append("INT");
        }
        else if (col.SqlType == "DECIMAL")
        {
            def.Append("DECIMAL(18,2)");
        }
        else if (col.SqlType == "MONEY")
        {
            def.Append("MONEY");
        }
        else // NVARCHAR or default
        {
            int length = (col.MaxLength.HasValue && col.MaxLength > 0) ? col.MaxLength.Value : 255;
            // Cap length at reasonable limit, use MAX for very long values
            if (length > 4000)
                def.Append("NVARCHAR(MAX)");
            else
                def.Append($"NVARCHAR({length})");
        }

        // Add NULL constraint
        if (!col.AllowNull && !col.IsPrimaryKey)
            def.Append(" NOT NULL");
        else if (col.AllowNull && !col.IsPrimaryKey)
            def.Append(" NULL");

        return def.ToString();
    }

    /// <summary>
    /// Generate INSERT INTO statements (batched in groups of 1000 to avoid SQL Server limits)
    /// </summary>
    private static string GenerateInsertStatements(DataTableSchema schema, string tableName, string schema_name, bool isTemporaryTable = false)
    {
        var sb = new StringBuilder();

        string fullTableName;
        if (isTemporaryTable)
        {
            fullTableName = $"[#{tableName}_Temporal]";
        }
        else
        {
            fullTableName = $"[{schema_name}].[{tableName}_Temporal]";
        }
        
        // Build column list (exclude temporal columns and IDENTITY PK if applicable)
        var insertColumns = new List<string>();
        for (int i = 0; i < schema.Columns.Count; i++)
        {
            var col = schema.Columns[i];
            // Skip IDENTITY columns (they auto-generate)
            if (col.IsPrimaryKey && col.SqlType == "INT")
                continue;
            insertColumns.Add($"[{col.ColumnName}]");
        }

        var columnList = string.Join(", ", insertColumns);

        const int BATCH_SIZE = 1000; // SQL Server limit for VALUES clause
        int totalRows = schema.DataRows.Count;
        int batchCount = (totalRows + BATCH_SIZE - 1) / BATCH_SIZE; // Ceiling division

        for (int batchIndex = 0; batchIndex < batchCount; batchIndex++)
        {
            int batchStart = batchIndex * BATCH_SIZE;
            int batchEnd = Math.Min(batchStart + BATCH_SIZE, totalRows);
            int batchRowCount = batchEnd - batchStart;

            // Start new INSERT for this batch
            sb.AppendLine($"INSERT INTO {fullTableName} ({columnList})");
            sb.AppendLine("VALUES");

            // Generate VALUES for this batch
            var batchValues = new List<string>();
            for (int rowIndex = batchStart; rowIndex < batchEnd; rowIndex++)
            {
                var row = schema.DataRows[rowIndex];
                var values = new List<string>();

                for (int colIndex = 0; colIndex < schema.Columns.Count; colIndex++)
                {
                    var col = schema.Columns[colIndex];
                    var value = row[colIndex];

                    // Skip IDENTITY columns
                    if (col.IsPrimaryKey && col.SqlType == "INT")
                        continue;

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
    /// Format a value for SQL based on its type
    /// Only handles basic types: INT, DECIMAL, MONEY, NVARCHAR
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
            "INT" => int.TryParse(value, out var intVal) ? intVal.ToString() : "NULL",
            "DECIMAL" => decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var decVal) ? decVal.ToString(CultureInfo.InvariantCulture) : "NULL",
            "MONEY" => decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var moneyVal) ? moneyVal.ToString(CultureInfo.InvariantCulture) : "NULL",
            _ => $"'{EscapeSqlString(value)}'"  // NVARCHAR or default
        };
    }

    /// <summary>
    /// Format boolean value for SQL
    /// </summary>
    private static string FormatBoolValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "NULL";

        bool isBool = bool.TryParse(value, out var boolVal);
        if (isBool)
            return boolVal ? "1" : "0";

        if (value == "1")
            return "1";
        if (value == "0")
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

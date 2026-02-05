using CopyAsInsert.Models;
using System.Globalization;

namespace CopyAsInsert.Services;

/// <summary>
/// Infers column types from data with lenient matching (70% threshold)
/// Auto-detects primary key from first INT or ID-named column
/// </summary>
public class TypeInferenceEngine
{
    private const int LENIENT_THRESHOLD = 70; // 70% of values must match the type

    /// <summary>
    /// Infer types for all columns in the schema
    /// </summary>
    public static void InferColumnTypes(DataTableSchema schema)
    {
        if (schema.Columns.Count == 0 || schema.DataRows.Count == 0)
            return;

        var pkIndex = -1;

        for (int colIndex = 0; colIndex < schema.Columns.Count; colIndex++)
        {
            var column = schema.Columns[colIndex];
            var values = schema.DataRows.Select(row => row[colIndex]).ToList();

            // Infer type based on values
            var inferredType = InferColumnType(values);
            column.SqlType = inferredType;
            column.ConfidencePercent = CalculateTypeConfidence(values, inferredType);

            // Set sample value
            var nonEmptyValue = values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
            column.SampleValue = nonEmptyValue ?? values.FirstOrDefault() ?? "";

            // Check for NULL values
            column.AllowNull = values.Any(v => string.IsNullOrWhiteSpace(v));

            // Set max length for VARCHAR
            if (inferredType == "VARCHAR")
            {
                column.MaxLength = values.Where(v => !string.IsNullOrWhiteSpace(v))
                    .Max(v => v.Length);
            }

            // Auto-detect primary key: first INT column or column with "ID" in name
            if (pkIndex == -1 && (inferredType == "INT" || column.ColumnName.Contains("ID", StringComparison.OrdinalIgnoreCase)))
            {
                pkIndex = colIndex;
                column.IsPrimaryKey = true;
                column.AllowNull = false;
            }
        }

        schema.PrimaryKeyColumnIndex = pkIndex;
    }

    /// <summary>
    /// Infer the type of a column based on its values (lenient matching: 70% threshold)
    /// Type precedence: INT → FLOAT → DATETIME2 → BOOL → VARCHAR
    /// </summary>
    private static string InferColumnType(List<string> values)
    {
        if (values.Count == 0)
            return "VARCHAR";

        var nonEmptyValues = values.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();

        if (nonEmptyValues.Count == 0)
            return "VARCHAR";

        // Try INT
        int intCount = nonEmptyValues.Count(v => int.TryParse(v, out _));
        if (IntPercentage(intCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "INT";

        // Try FLOAT
        int floatCount = nonEmptyValues.Count(v => double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _));
        if (PercentageMatch(floatCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "FLOAT";

        // Try DATETIME2
        int dateCount = nonEmptyValues.Count(v => DateTime.TryParse(v, out _));
        if (PercentageMatch(dateCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "DATETIME2";

        // Try BOOL
        int boolCount = nonEmptyValues.Count(v => v.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || 
                                                    v.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                                                    v == "0" || v == "1");
        if (PercentageMatch(boolCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "BIT";

        // Default to VARCHAR
        return "VARCHAR";
    }

    /// <summary>
    /// Calculate confidence percentage for a type match
    /// </summary>
    private static int CalculateTypeConfidence(List<string> values, string sqlType)
    {
        if (values.Count == 0)
            return 0;

        var nonEmptyValues = values.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();

        if (nonEmptyValues.Count == 0)
            return 0;

        int matchCount = sqlType switch
        {
            "INT" => nonEmptyValues.Count(v => int.TryParse(v, out _)),
            "FLOAT" => nonEmptyValues.Count(v => double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _)),
            "DATETIME2" => nonEmptyValues.Count(v => DateTime.TryParse(v, out _)),
            "BIT" => nonEmptyValues.Count(v => v.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || 
                                                 v.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                                                 v == "0" || v == "1"),
            _ => nonEmptyValues.Count
        };

        return (int)((matchCount / (double)nonEmptyValues.Count) * 100);
    }

    private static int IntPercentage(int matchCount, int totalCount)
    {
        return (int)((matchCount / (double)totalCount) * 100);
    }

    private static int PercentageMatch(int matchCount, int totalCount)
    {
        return (int)((matchCount / (double)totalCount) * 100);
    }
}

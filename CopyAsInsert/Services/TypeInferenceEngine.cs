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

            // All columns are nullable by default
            column.AllowNull = true;

            // Set max length for NVARCHAR
            if (inferredType == "NVARCHAR")
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
    /// Only uses basic types: INT, DECIMAL, MONEY, NVARCHAR
    /// If uncertain, defaults to NVARCHAR for robustness
    /// Type precedence: INT → DECIMAL → NVARCHAR
    /// </summary>
    private static string InferColumnType(List<string> values)
    {
        if (values.Count == 0)
            return "NVARCHAR";

        var nonEmptyValues = values.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();

        if (nonEmptyValues.Count == 0)
            return "NVARCHAR";

        // Try INT - must be strict integers
        int intCount = nonEmptyValues.Count(v => int.TryParse(v, out _));
        if (IntPercentage(intCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "INT";

        // Try DECIMAL - includes floats, decimals
        int decimalCount = nonEmptyValues.Count(v => decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _));
        if (PercentageMatch(decimalCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "DECIMAL";

        // Default to NVARCHAR - when in doubt, use text for robustness
        return "NVARCHAR";
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
            "DECIMAL" => nonEmptyValues.Count(v => decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _)),
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

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
                var nonEmptyValueLengths = values.Where(v => !string.IsNullOrWhiteSpace(v))
                    .Select(v => v.Length)
                    .ToList();
                
                column.MaxLength = nonEmptyValueLengths.Count > 0 ? nonEmptyValueLengths.Max() : 0;
            }
        }
    }

    /// <summary>
    /// Infer the type of a column based on its values
    /// Supports: INT, MONEY (decimal with 2-4 decimals), NVARCHAR
    /// If any value is "0" or has leading zeros, treat as NVARCHAR
    /// </summary>
    private static string InferColumnType(List<string> values)
    {
        if (values.Count == 0)
            return "NVARCHAR";

        var nonEmptyValues = values.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();

        if (nonEmptyValues.Count == 0)
            return "NVARCHAR";

        // If any value has leading zeros (like "0001", "0010") or is exactly "0", treat as NVARCHAR
        if (nonEmptyValues.Any(v => 
            (v.StartsWith("0") && v.Length > 1) || 
            v == "0"))
            return "NVARCHAR";

        // Try INT - must be strict integers (70% threshold)
        int intCount = nonEmptyValues.Count(v => int.TryParse(v, out _));
        if (IntPercentage(intCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "INT";

        // Try MONEY - decimal values with reasonable precision
        int moneyCount = nonEmptyValues.Count(v => 
            decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _));
        if (PercentageMatch(moneyCount, nonEmptyValues.Count) >= LENIENT_THRESHOLD)
            return "MONEY";

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
            "MONEY" => nonEmptyValues.Count(v => decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _)),
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

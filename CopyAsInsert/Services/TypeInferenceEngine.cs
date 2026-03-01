using CopyAsInsert.Models;
using System.Globalization;
using System.Text.RegularExpressions;

namespace CopyAsInsert.Services;

/// <summary>
/// Professional type inference engine with support for:
/// INT, FLOAT, DATETIME2, NVARCHAR
/// Uses confidence thresholds: INT (85%), FLOAT (85%), DATETIME (80%), NVARCHAR (fallback)
/// </summary>
public class TypeInferenceEngine
{
    // Confidence thresholds for each type (stricter = safer fallback to NVARCHAR)
    private const double INT_THRESHOLD = 0.85;           // 85% of values must be valid integers
    private const double FLOAT_THRESHOLD = 0.85;         // 85% of values must be valid decimals (excluding pure ints)
    private const double DATETIME_THRESHOLD = 0.80;      // 80% of values must match date patterns

    /// Sets SqlType, ConfidencePercent, ConfidenceScore, and InferenceReason for each column
    /// </summary>
    public static void InferColumnTypes(DataTableSchema schema)
    {
        if (schema.Columns.Count == 0 || schema.DataRows.Count == 0)
            return;

        for (int colIndex = 0; colIndex < schema.Columns.Count; colIndex++)
        {
            var column = schema.Columns[colIndex];
            var values = schema.DataRows.Select(row => row[colIndex]).ToList();

            // Analyze the column
            var (inferredType, confidence, reason) = InferColumnType(values);
            column.SqlType = inferredType;
            column.ConfidencePercent = (int)(confidence * 100);
            column.ConfidenceScore = confidence;
            column.InferenceReason = reason;

            // Set sample value
            var nonEmptyValue = values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v) && !IsNullRepresentation(v));
            column.SampleValue = nonEmptyValue ?? values.FirstOrDefault() ?? "";

            // Determine if column allows NULL (check for empty/null values and common NULL representations)
            column.AllowNull = values.Any(v => IsNullRepresentation(v));

            // Set max length for NVARCHAR
            if (inferredType == "NVARCHAR")
            {
                var nonEmptyValueLengths = values.Where(v => !string.IsNullOrWhiteSpace(v) && !IsNullRepresentation(v))
                    .Select(v => v.Length)
                    .ToList();
                
                column.MaxLength = nonEmptyValueLengths.Count > 0 ? nonEmptyValueLengths.Max() : 0;
            }

            // Log inference decision
            Logger.LogDebug($"Column '{column.ColumnName}': {inferredType} ({column.ConfidencePercent}%) - {reason}");
        }
    }

    /// <summary>
    /// Infer the type of a column based on its values
    /// Returns tuple of (TypeName, Confidence, Reason)
    /// </summary>
    private static (string type, double confidence, string reason) InferColumnType(List<string> values)
    {
        if (values.Count == 0)
            return ("NVARCHAR", 1.0, "No data provided");

        // Filter out empty/whitespace-only values for analysis
        var nonEmptyValues = values.Where(v => !IsNullRepresentation(v))
                                   .ToList();

        if (nonEmptyValues.Count == 0)
            return ("NVARCHAR", 1.0, "All values are empty or NULL");

        // Try each type in priority order, using thresholds
        // Note: BIT type is disabled to preserve data integrity
        
        // 1. Try DATETIME2 - moderate threshold (80%) due to format variability
        var (dateCount, dateConfidence) = DetectDateTime(nonEmptyValues);
        if (dateConfidence >= DATETIME_THRESHOLD)
            return ("DATETIME2", dateConfidence, $"{(int)(dateConfidence * 100)}% match date patterns (ISO/US/EU formats)");

        // 2. Try INT - strict threshold (85%)
        var (intCount, intConfidence) = DetectInteger(nonEmptyValues);
        if (intConfidence >= INT_THRESHOLD)
            return ("INT", intConfidence, $"{(int)(intConfidence * 100)}% are valid integers");

        // 3. Try FLOAT - strict threshold (85%), excluding pure integers
        var (floatCount, floatConfidence) = DetectFloat(nonEmptyValues);
        if (floatConfidence >= FLOAT_THRESHOLD)
            return ("FLOAT", floatConfidence, $"{(int)(floatConfidence * 100)}% are valid decimals");

        // 5. Default to NVARCHAR
        return ("NVARCHAR", 1.0, "No pattern matched; defaulting to text");
    }

    /// <summary>
    /// Detect if values are valid integers
    /// Rejects values with leading zeros (e.g., "0001", "0123") to preserve them as NVARCHAR
    /// Handles European number format with thousand separators
    /// </summary>
    private static (int matchCount, double confidence) DetectInteger(List<string> values)
    {
        int matchCount = 0;
        foreach (var value in values)
        {
            // Reject values with decimal separators (period or comma)
            if (value.Contains('.') || value.Contains(','))
            {
                continue;
            }

            // Reject values with leading zeros (except "0" itself) to preserve source data
            // Examples: "0001", "0123", "00456" are NVARCHAR, not INT
            string trimmed = value.Trim();
            if (trimmed.Length > 1 && trimmed[0] == '0' && char.IsDigit(trimmed[1]))
            {
                continue; // Has leading zero, treat as NVARCHAR
            }

            // Try to parse as Int64 to handle larger integers
            if (long.TryParse(value, out _))
            {
                matchCount++;
            }
        }

        double confidence = values.Count > 0 ? (double)matchCount / values.Count : 0;
        return (matchCount, confidence);
    }

    /// <summary>
    /// Detect if values are valid floating-point numbers
    /// Handles both period (.) and comma (,) as decimal separators
    /// Excludes pure integers (which should be INT, not FLOAT)
    /// </summary>
    private static (int matchCount, double confidence) DetectFloat(List<string> values)
    {
        int matchCount = 0;
        foreach (var value in values)
        {
            // Check if it's a valid decimal with period separator
            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalVal))
            {
                // Include only if it actually has a fractional part (e.g., 1.5, not 1.0)
                string normalized = decimalVal.ToString(CultureInfo.InvariantCulture);
                if (value.Contains('.') || value.Contains(',') || normalized.Contains('.'))
                {
                    matchCount++;
                    continue;
                }
            }

            // Check if it's a valid decimal with comma separator (European format: 1,5 or 1.234,56)
            if (value.Contains(','))
            {
                string normalizedForComma = value.Replace(".", "").Replace(",", ".");
                if (decimal.TryParse(normalizedForComma, NumberStyles.Any, CultureInfo.InvariantCulture, out var commaDecimalVal))
                {
                    // If it has a comma, it's definitely a decimal (European format)
                    matchCount++;
                }
            }
        }

        double confidence = values.Count > 0 ? (double)matchCount / values.Count : 0;
        return (matchCount, confidence);
    }

    /// <summary>
    /// Detect if values are dates/datetimes
    /// Supports multiple formats: ISO-8601, US (MM/DD/YYYY), EU (DD/MM/YYYY), and variations
    /// </summary>
    private static (int matchCount, double confidence) DetectDateTime(List<string> values)
    {
        int matchCount = 0;
        var datePatterns = new[]
        {
            // ISO 8601 formats
            @"^\d{4}-\d{2}-\d{2}(T|\s)\d{2}:\d{2}:\d{2}",     // 2026-02-06T15:30:00 or 2026-02-06 15:30:00
            @"^\d{4}-\d{2}-\d{2}$",                            // 2026-02-06
            
            // US format (MM/DD/YYYY or MM-DD-YYYY)
            @"^(0?[1-9]|1[0-2])[/-](0?[1-9]|[12]\d|3[01])[/-]\d{4}$",
            @"^(0?[1-9]|1[0-2])[/-](0?[1-9]|[12]\d|3[01])[/-](\d{2}|\d{4})$",
            
            // EU format (DD/MM/YYYY or DD-MM-YYYY)
            @"^(0?[1-9]|[12]\d|3[01])[/-](0?[1-9]|1[0-2])[/-]\d{4}$",
            @"^(0?[1-9]|[12]\d|3[01])[/-](0?[1-9]|1[0-2])[/-](\d{2}|\d{4})$",
            
            // Other variations
            @"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2},? \d{4}",  // January 6, 2026
            @"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*, ",                                     // Monday, ...
        };

        foreach (var value in values)
        {
            // Try parsing with common .NET patterns
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _) ||
                DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out _))
            {
                matchCount++;
                continue;
            }

            // Check regex patterns
            if (datePatterns.Any(pattern => Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase)))
            {
                matchCount++;
            }
        }

        double confidence = values.Count > 0 ? (double)matchCount / values.Count : 0;
        return (matchCount, confidence);
    }

    /// <summary>
    /// Check if a value represents NULL (empty string, "NULL" keyword, or other representations)
    /// </summary>
    private static bool IsNullRepresentation(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return true;

        // Check for explicit NULL keyword (case-insensitive)
        if (value.Equals("NULL", StringComparison.OrdinalIgnoreCase))
            return true;

        // Check for variants with different cases
        string trimmed = value.Trim();
        if (trimmed.Equals("NULL", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("Null", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("null", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("#N/A", StringComparison.OrdinalIgnoreCase) ||  // Excel error
            trimmed.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("<null>", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("(null)", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }
}

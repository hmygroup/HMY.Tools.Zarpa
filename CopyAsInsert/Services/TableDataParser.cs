using CopyAsInsert.Models;
using ClosedXML.Excel;
using System.Globalization;
using System.Text;

namespace CopyAsInsert.Services;

/// <summary>
/// Parses tabular data from clipboard (TSV/CSV) and Excel files
/// </summary>
public class TableDataParser
{
    /// <summary>
    /// Normalize column names by removing accents and spaces
    /// </summary>
    private static string NormalizeColumnName(string columnName)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            return columnName;

        // Decompose accented characters (á -> a + accent mark)
        string normalized = columnName.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();

        // Remove non-spacing mark characters (accents)
        foreach (var c in normalized)
        {
            var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
            if (unicodeCategory != UnicodeCategory.NonSpacingMark)
            {
                sb.Append(c);
            }
        }

        // Remove spaces
        normalized = sb.ToString().Replace(" ", "", StringComparison.OrdinalIgnoreCase);

        // Recompose to normal form
        return normalized.Normalize(NormalizationForm.FormC);
    }
    /// <summary>
    /// Parse clipboard text as TSV or CSV (also supports single values)
    /// </summary>
    public static DataTableSchema ParseClipboardText(string clipboardText, bool hasHeaders = true)
    {
        if (string.IsNullOrWhiteSpace(clipboardText))
            throw new InvalidOperationException("Clipboard is empty");

        var lines = clipboardText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        if (lines.Length < 1)
            throw new InvalidOperationException("No data found in clipboard");

        // Allow single row if no headers
        if (hasHeaders && lines.Length < 1)
            throw new InvalidOperationException("At least one header row required");

        // Detect delimiter (tab preference, fallback to comma, fallback to no delimiter)
        char delimiter = '\t';
        if (!lines[0].Contains('\t'))
        {
            delimiter = ',';
            if (!lines[0].Contains(','))
            {
                delimiter = '\0'; // No delimiter found - treat as single column
            }
        }

        var schema = new DataTableSchema
        {
            OriginalTableName = "ParsedTable",
            DataSource = delimiter == '\t' ? "ClipboardTSV" : (delimiter == ',' ? "ClipboardCSV" : "ClipboardSingle")
        };

        string[] headers;
        int dataStartIndex;

        if (hasHeaders)
        {
            // First row is headers
            headers = delimiter == '\0' 
                ? new[] { NormalizeColumnName(lines[0].Trim()) }
                : lines[0].Split(delimiter).Select(h => NormalizeColumnName(h.Trim())).ToArray();
            dataStartIndex = 1;

            if (headers.Length == 0)
                throw new InvalidOperationException("No columns found in header");
        }
        else
        {
            // No headers - generate generic column names based on first row
            var firstRow = delimiter == '\0'
                ? new[] { lines[0].Trim() }
                : lines[0].Split(delimiter);
            headers = new string[firstRow.Length];
            for (int i = 0; i < firstRow.Length; i++)
            {
                headers[i] = $"Col{i + 1}";
            }
            dataStartIndex = 0; // All rows are data
        }

        // Add column headers
        foreach (var header in headers)
        {
            schema.Columns.Add(new ColumnTypeInfo 
            { 
                ColumnName = header,
                SqlType = "VARCHAR",
                AllowNull = true
            });
        }

        // Parse data rows
        for (int i = dataStartIndex; i < lines.Length; i++)
        {
            string[] values;
            if (delimiter == '\0')
            {
                // No delimiter - single column
                values = new[] { lines[i].Trim() };
            }
            else
            {
                values = lines[i].Split(delimiter);
            }
            
            // Pad with empty strings if columns mismatch
            var paddedValues = new List<string>();
            for (int j = 0; j < headers.Length; j++)
            {
                paddedValues.Add(j < values.Length ? values[j].Trim() : string.Empty);
            }

            schema.DataRows.Add(paddedValues.ToArray());
        }

        return schema;
    }

    /// <summary>
    /// Parse an Excel file (.xlsx)
    /// </summary>
    public static DataTableSchema ParseExcelFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Only .xlsx files are supported");

        var workbook = new XLWorkbook(filePath);
        
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet == null)
            throw new InvalidOperationException("No worksheets found in Excel file");

        var usedRange = worksheet.RangeUsed();
        if (usedRange == null)
            throw new InvalidOperationException("No data found in worksheet");

        var schema = new DataTableSchema
        {
            OriginalTableName = Path.GetFileNameWithoutExtension(filePath),
            DataSource = "ExcelFile"
        };

        var rows = usedRange.Rows().ToList();

        if (rows.Count < 2)
            throw new InvalidOperationException("At least one header row and one data row required");

        // Parse headers
        var headerRow = rows[0];
        var headers = new List<string>();
        foreach (var cell in headerRow.Cells())
        {
            var headerName = cell.Value?.ToString()?.Trim() ?? "";
            headers.Add(NormalizeColumnName(headerName));
        }

        // Add column info
        foreach (var header in headers)
        {
            schema.Columns.Add(new ColumnTypeInfo
            {
                ColumnName = header,
                SqlType = "VARCHAR",
                AllowNull = true
            });
        }

        // Parse data rows
        for (int i = 1; i < rows.Count; i++)
        {
            var dataRow = rows[i];
            var values = new List<string>();

            foreach (var cell in dataRow.Cells().Take(headers.Count))
            {
                values.Add(cell.Value?.ToString() ?? string.Empty);
            }

            schema.DataRows.Add(values.ToArray());
        }

        return schema;
    }

    /// <summary>
    /// Detect if text from clipboard appears to be tabular (TSV/CSV)
    /// </summary>
    public static bool IsValidTabularText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2)
            return false;

        var firstLine = lines[0];
        bool hasDelimiter = firstLine.Contains('\t') || firstLine.Contains(',');

        return hasDelimiter;
    }
}

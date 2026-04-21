using System.Text.RegularExpressions;

namespace CopyAsInsert.Services;

internal sealed class QueryImportDefinition
{
    public string Script { get; init; } = string.Empty;
    public string SuggestedName { get; init; } = "Result";
}

internal static class SqlImportQueryPlanner
{
    private static readonly Regex SourceNameRegex = new(
        @"\bFROM\s+(?<source>(?:\[[^\]]+\]|[#@A-Za-z0-9_]+)(?:\.(?:\[[^\]]+\]|[#@A-Za-z0-9_]+))*)",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public static List<QueryImportDefinition> BuildImportQueries(string script)
    {
        string normalizedScript = script?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalizedScript))
        {
            return new List<QueryImportDefinition>
            {
                new() { Script = string.Empty, SuggestedName = "Result_1" }
            };
        }

        List<SqlStatementBlock> statements = SqlScriptTokenizer.SplitTopLevelStatements(normalizedScript);
        if (statements.Count == 0)
        {
            return new List<QueryImportDefinition>
            {
                CreateQueryImport(normalizedScript, 1)
            };
        }

        int firstResultIndex = FindFirstTrailingResultIndex(statements);
        int trailingResultCount = statements.Count - firstResultIndex;
        if (trailingResultCount <= 1)
        {
            return new List<QueryImportDefinition>
            {
                CreateQueryImport(normalizedScript, 1)
            };
        }

        string prefix = string.Join(
            Environment.NewLine + Environment.NewLine,
            statements.Take(firstResultIndex).Select(statement => statement.Text.Trim()).Where(statement => !string.IsNullOrWhiteSpace(statement)));

        List<QueryImportDefinition> results = new();
        for (int index = firstResultIndex; index < statements.Count; index++)
        {
            string resultStatement = statements[index].Text.Trim();
            string importScript = string.IsNullOrWhiteSpace(prefix)
                ? resultStatement
                : $"{prefix}{Environment.NewLine}{Environment.NewLine}{resultStatement}";

            results.Add(new QueryImportDefinition
            {
                Script = importScript,
                SuggestedName = GetSuggestedQueryName(resultStatement, results.Count + 1)
            });
        }

        return results;
    }

    private static QueryImportDefinition CreateQueryImport(string script, int ordinal)
    {
        return new QueryImportDefinition
        {
            Script = script,
            SuggestedName = GetSuggestedQueryName(script, ordinal)
        };
    }

    private static int FindFirstTrailingResultIndex(IReadOnlyList<SqlStatementBlock> statements)
    {
        int firstResultIndex = statements.Count;
        for (int index = statements.Count - 1; index >= 0; index--)
        {
            if (IsExportResultStatement(statements[index]))
            {
                firstResultIndex = index;
                continue;
            }

            break;
        }

        return firstResultIndex;
    }

    private static bool IsExportResultStatement(SqlStatementBlock statement)
    {
        return string.Equals(statement.CommandKind, "SELECT", StringComparison.OrdinalIgnoreCase)
            && !SqlScriptTokenizer.GetTopLevelTokens(statement.Text).Contains("INTO", StringComparer.OrdinalIgnoreCase);
    }

    private static string GetSuggestedQueryName(string statement, int ordinal)
    {
        Match match = SourceNameRegex.Match(statement);
        if (match.Success)
        {
            string rawSource = match.Groups["source"].Value;
            string[] parts = rawSource.Split('.', StringSplitOptions.RemoveEmptyEntries);
            string leafName = parts.Length == 0 ? rawSource : parts[^1];
            leafName = leafName.Trim().Trim('[', ']').TrimStart('#', '@');
            string sanitized = SanitizeSuggestedName(leafName);
            if (!string.IsNullOrWhiteSpace(sanitized))
            {
                return sanitized;
            }
        }

        return $"Result_{ordinal}";
    }

    private static string SanitizeSuggestedName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Result";
        }

        string sanitized = Regex.Replace(value.Trim(), @"[^A-Za-z0-9_]+", "_").Trim('_');
        return string.IsNullOrWhiteSpace(sanitized) ? "Result" : sanitized;
    }
}
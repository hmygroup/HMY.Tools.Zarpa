using System.Text;

namespace CopyAsInsert.Services;

internal sealed class SqlStatementBlock
{
    public string Text { get; init; } = string.Empty;
    public string? CommandKind { get; init; }
}

internal static class SqlScriptTokenizer
{
    private static readonly HashSet<string> StatementStartKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "SELECT",
        "WITH",
        "INSERT",
        "UPDATE",
        "DELETE",
        "MERGE",
        "DROP",
        "CREATE",
        "ALTER",
        "DECLARE",
        "SET",
        "EXEC",
        "EXECUTE",
        "TRUNCATE",
        "IF",
        "BEGIN",
        "END",
        "PRINT",
        "RAISERROR",
        "RETURN"
    };

    public static List<SqlStatementBlock> SplitTopLevelStatements(string script)
    {
        List<SqlStatementBlock> statements = new();
        if (string.IsNullOrWhiteSpace(script))
        {
            return statements;
        }

        string normalizedScript = script.Replace("\r\n", "\n");
        using StringReader reader = new(normalizedScript);
        StringBuilder currentStatement = new();
        string? currentStatementKeyword = null;
        string? lastTopLevelKeyword = null;
        int parenthesisDepth = 0;
        bool inBlockComment = false;
        bool inStringLiteral = false;
        bool inBracketIdentifier = false;

        while (reader.ReadLine() is string line)
        {
            string? lineKeyword = GetLineLeadingKeyword(line, inBlockComment, inStringLiteral, inBracketIdentifier, parenthesisDepth);

            if (lineKeyword != null && ShouldStartNewStatement(lineKeyword, currentStatementKeyword, lastTopLevelKeyword, currentStatement.Length > 0))
            {
                AppendStatementBlock(statements, currentStatement, currentStatementKeyword);
                currentStatement.Clear();
                currentStatementKeyword = lineKeyword;
                lastTopLevelKeyword = lineKeyword;
            }
            else if (currentStatement.Length == 0 && lineKeyword != null)
            {
                currentStatementKeyword = lineKeyword;
                lastTopLevelKeyword = lineKeyword;
            }

            if (currentStatement.Length > 0 || !string.IsNullOrWhiteSpace(line))
            {
                currentStatement.AppendLine(line);
            }

            UpdateParserState(line + "\n", ref parenthesisDepth, ref inBlockComment, ref inStringLiteral, ref inBracketIdentifier, ref lastTopLevelKeyword);
        }

        AppendStatementBlock(statements, currentStatement, currentStatementKeyword);
        return statements;
    }

    public static List<string> GetTopLevelTokens(string sql)
    {
        List<string> tokens = new();
        int parenthesisDepth = 0;
        bool inBlockComment = false;
        bool inLineComment = false;
        bool inStringLiteral = false;
        bool inBracketIdentifier = false;

        for (int index = 0; index < sql.Length; index++)
        {
            char current = sql[index];
            char next = index + 1 < sql.Length ? sql[index + 1] : '\0';

            if (inLineComment)
            {
                if (current == '\n')
                {
                    inLineComment = false;
                }

                continue;
            }

            if (inBlockComment)
            {
                if (current == '*' && next == '/')
                {
                    inBlockComment = false;
                    index++;
                }

                continue;
            }

            if (inStringLiteral)
            {
                if (current == '\'' && next == '\'')
                {
                    index++;
                    continue;
                }

                if (current == '\'')
                {
                    inStringLiteral = false;
                }

                continue;
            }

            if (inBracketIdentifier)
            {
                if (current == ']')
                {
                    inBracketIdentifier = false;
                }

                continue;
            }

            if (current == '-' && next == '-')
            {
                inLineComment = true;
                index++;
                continue;
            }

            if (current == '/' && next == '*')
            {
                inBlockComment = true;
                index++;
                continue;
            }

            if (current == '\'')
            {
                inStringLiteral = true;
                continue;
            }

            if (current == '[')
            {
                inBracketIdentifier = true;
                continue;
            }

            if (current == '(')
            {
                parenthesisDepth++;
                continue;
            }

            if (current == ')')
            {
                parenthesisDepth = Math.Max(0, parenthesisDepth - 1);
                continue;
            }

            if (parenthesisDepth != 0 || !char.IsLetter(current))
            {
                continue;
            }

            int start = index;
            while (index + 1 < sql.Length && (char.IsLetter(sql[index + 1]) || sql[index + 1] == '_'))
            {
                index++;
            }

            tokens.Add(sql[start..(index + 1)].ToUpperInvariant());
        }

        return tokens;
    }

    private static void AppendStatementBlock(List<SqlStatementBlock> statements, StringBuilder currentStatement, string? currentStatementKeyword)
    {
        string statementText = currentStatement.ToString().Trim();
        if (string.IsNullOrWhiteSpace(statementText))
        {
            return;
        }

        string? commandKind = currentStatementKeyword ?? GetStatementCommandKind(statementText);
        if (commandKind == null)
        {
            return;
        }

        statements.Add(new SqlStatementBlock
        {
            Text = statementText,
            CommandKind = commandKind
        });
    }

    private static string? GetLineLeadingKeyword(string line, bool inBlockComment, bool inStringLiteral, bool inBracketIdentifier, int parenthesisDepth)
    {
        if (inBlockComment || inStringLiteral || inBracketIdentifier || parenthesisDepth != 0)
        {
            return null;
        }

        int index = 0;
        while (index < line.Length && char.IsWhiteSpace(line[index]))
        {
            index++;
        }

        if (index >= line.Length || (line[index] == '-' && index + 1 < line.Length && line[index + 1] == '-') || (line[index] == '/' && index + 1 < line.Length && line[index + 1] == '*'))
        {
            return null;
        }

        if (!char.IsLetter(line[index]))
        {
            return null;
        }

        int start = index;
        while (index < line.Length && (char.IsLetter(line[index]) || line[index] == '_'))
        {
            index++;
        }

        string keyword = line[start..index].ToUpperInvariant();
        return IsStatementStartKeyword(keyword) ? keyword : null;
    }

    private static bool ShouldStartNewStatement(string lineKeyword, string? currentStatementKeyword, string? lastTopLevelKeyword, bool hasCurrentStatement)
    {
        if (!hasCurrentStatement)
        {
            return true;
        }

        if (!IsStatementStartKeyword(lineKeyword))
        {
            return false;
        }

        if (string.Equals(lineKeyword, "SELECT", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(currentStatementKeyword, "INSERT", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(currentStatementKeyword, "WITH", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(currentStatementKeyword, "CREATE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(currentStatementKeyword, "ALTER", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lastTopLevelKeyword, "UNION", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lastTopLevelKeyword, "ALL", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lastTopLevelKeyword, "EXCEPT", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lastTopLevelKeyword, "INTERSECT", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lastTopLevelKeyword, "AS", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsStatementStartKeyword(string keyword)
    {
        return StatementStartKeywords.Contains(keyword);
    }

    private static void UpdateParserState(
        string sqlFragment,
        ref int parenthesisDepth,
        ref bool inBlockComment,
        ref bool inStringLiteral,
        ref bool inBracketIdentifier,
        ref string? lastTopLevelKeyword)
    {
        bool inLineComment = false;

        for (int index = 0; index < sqlFragment.Length; index++)
        {
            char current = sqlFragment[index];
            char next = index + 1 < sqlFragment.Length ? sqlFragment[index + 1] : '\0';

            if (inLineComment)
            {
                if (current == '\n')
                {
                    inLineComment = false;
                }

                continue;
            }

            if (inBlockComment)
            {
                if (current == '*' && next == '/')
                {
                    inBlockComment = false;
                    index++;
                }

                continue;
            }

            if (inStringLiteral)
            {
                if (current == '\'' && next == '\'')
                {
                    index++;
                    continue;
                }

                if (current == '\'')
                {
                    inStringLiteral = false;
                }

                continue;
            }

            if (inBracketIdentifier)
            {
                if (current == ']')
                {
                    inBracketIdentifier = false;
                }

                continue;
            }

            if (current == '-' && next == '-')
            {
                inLineComment = true;
                index++;
                continue;
            }

            if (current == '/' && next == '*')
            {
                inBlockComment = true;
                index++;
                continue;
            }

            if (current == '\'')
            {
                inStringLiteral = true;
                continue;
            }

            if (current == '[')
            {
                inBracketIdentifier = true;
                continue;
            }

            if (current == '(')
            {
                parenthesisDepth++;
                continue;
            }

            if (current == ')')
            {
                parenthesisDepth = Math.Max(0, parenthesisDepth - 1);
                continue;
            }

            if (parenthesisDepth == 0 && char.IsLetter(current))
            {
                int start = index;
                while (index + 1 < sqlFragment.Length && (char.IsLetter(sqlFragment[index + 1]) || sqlFragment[index + 1] == '_'))
                {
                    index++;
                }

                lastTopLevelKeyword = sqlFragment[start..(index + 1)].ToUpperInvariant();
            }
        }
    }

    private static string? GetStatementCommandKind(string statement)
    {
        List<string> topLevelTokens = GetTopLevelTokens(statement);
        if (topLevelTokens.Count == 0)
        {
            return null;
        }

        if (!string.Equals(topLevelTokens[0], "WITH", StringComparison.OrdinalIgnoreCase))
        {
            return topLevelTokens[0];
        }

        foreach (string token in topLevelTokens)
        {
            if (string.Equals(token, "SELECT", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "INSERT", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "UPDATE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "DELETE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "MERGE", StringComparison.OrdinalIgnoreCase))
            {
                return token;
            }
        }

        return topLevelTokens[0];
    }
}
namespace DatabaseManager.Core.Services;

public readonly record struct QueryOutputModeParseResult(string Sql, bool FullOutputEnabled, bool HasFullDirective);

public static class QueryOutputModeParser
{
    public static QueryOutputModeParseResult Parse(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql))
        {
            return new QueryOutputModeParseResult(sql, false, false);
        }

        var lastNonWhitespace = FindLastNonWhitespaceIndex(sql);
        if (lastNonWhitespace < 0)
        {
            return new QueryOutputModeParseResult(sql, false, false);
        }

        var lineStart = sql.LastIndexOf('\n', lastNonWhitespace);
        lineStart = lineStart < 0 ? 0 : lineStart + 1;

        var line = sql[lineStart..(lastNonWhitespace + 1)];
        var commentStartInLine = FindLineCommentStartOutsideStringLiteral(line);
        if (commentStartInLine < 0)
        {
            return new QueryOutputModeParseResult(sql, false, false);
        }

        var commentText = line[(commentStartInLine + 2)..].Trim();
        if (!commentText.Equals("full", StringComparison.OrdinalIgnoreCase))
        {
            return new QueryOutputModeParseResult(sql, false, false);
        }

        var absoluteCommentStart = lineStart + commentStartInLine;
        var sanitizedSql = sql[..absoluteCommentStart].TrimEnd();
        return new QueryOutputModeParseResult(sanitizedSql, true, true);
    }

    public static IReadOnlyList<string> ExtractParameterNames(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql))
        {
            return Array.Empty<string>();
        }

        var names = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var insideString = false;
        var insideLineComment = false;
        var insideBlockComment = false;

        for (var i = 0; i < sql.Length; i++)
        {
            var current = sql[i];
            var next = i < sql.Length - 1 ? sql[i + 1] : '\0';

            if (insideLineComment)
            {
                if (current == '\n')
                {
                    insideLineComment = false;
                }

                continue;
            }

            if (insideBlockComment)
            {
                if (current == '*' && next == '/')
                {
                    insideBlockComment = false;
                    i++;
                }

                continue;
            }

            if (insideString)
            {
                if (current == '\'' && next == '\'')
                {
                    i++;
                    continue;
                }

                if (current == '\'')
                {
                    insideString = false;
                }

                continue;
            }

            if (current == '\'')
            {
                insideString = true;
                continue;
            }

            if (current == '-' && next == '-')
            {
                insideLineComment = true;
                i++;
                continue;
            }

            if (current == '/' && next == '*')
            {
                insideBlockComment = true;
                i++;
                continue;
            }

            if (current != '@')
            {
                continue;
            }

            if (next == '@')
            {
                i++;
                continue;
            }

            if (next != '_' && !char.IsLetter(next))
            {
                continue;
            }

            var start = i;
            var end = i + 1;
            while (end < sql.Length && (sql[end] == '_' || char.IsLetterOrDigit(sql[end])))
            {
                end++;
            }

            var parameterName = sql[start..end];
            if (seen.Add(parameterName))
            {
                names.Add(parameterName);
            }

            i = end - 1;
        }

        return names;
    }

    private static int FindLastNonWhitespaceIndex(string value)
    {
        for (var i = value.Length - 1; i >= 0; i--)
        {
            if (!char.IsWhiteSpace(value[i]))
            {
                return i;
            }
        }

        return -1;
    }

    private static int FindLineCommentStartOutsideStringLiteral(ReadOnlySpan<char> line)
    {
        var insideString = false;

        for (var i = 0; i < line.Length - 1; i++)
        {
            var current = line[i];
            var next = line[i + 1];

            if (current == '\'')
            {
                if (insideString && next == '\'')
                {
                    i++;
                    continue;
                }

                insideString = !insideString;
                continue;
            }

            if (!insideString && current == '-' && next == '-')
            {
                return i;
            }
        }

        return -1;
    }
}
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

    public static bool IsCreateOrAlterRoutineStatement(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql))
        {
            return false;
        }

        var span = SkipLeadingWhitespaceAndComments(sql.AsSpan());
        if (!TryConsumeKeyword(ref span, "CREATE"))
        {
            if (!TryConsumeKeyword(ref span, "ALTER"))
            {
                return false;
            }
        }
        else
        {
            if (TryConsumeKeyword(ref span, "OR"))
            {
                if (!TryConsumeKeyword(ref span, "ALTER"))
                {
                    return false;
                }
            }
        }

        return TryConsumeKeyword(ref span, "PROCEDURE")
            || TryConsumeKeyword(ref span, "PROC")
            || TryConsumeKeyword(ref span, "FUNCTION")
            || TryConsumeKeyword(ref span, "TRIGGER");
    }

    /// <summary>
    /// Locates the leading "CREATE [OR ALTER] PROC|PROCEDURE" keywords of a procedure definition,
    /// skipping leading whitespace/comments first, so callers can rewrite just those keywords.
    /// Matching only at the statement start (not anywhere in the text) keeps a "CREATE PROCEDURE"
    /// mentioned in a header comment or string literal from being mistaken for the real header.
    /// </summary>
    public static bool TryFindCreateProcedureHeader(string sql, out int start, out int length)
    {
        start = 0;
        length = 0;
        if (string.IsNullOrWhiteSpace(sql))
        {
            return false;
        }

        var span = SkipLeadingWhitespaceAndComments(sql.AsSpan());
        var headerStart = sql.Length - span.Length;

        if (!TryConsumeKeyword(ref span, "CREATE"))
        {
            return false;
        }

        if (TryConsumeKeyword(ref span, "OR") && !TryConsumeKeyword(ref span, "ALTER"))
        {
            return false;
        }

        if (!TryConsumeKeyword(ref span, "PROCEDURE") && !TryConsumeKeyword(ref span, "PROC"))
        {
            return false;
        }

        start = headerStart;
        length = sql.Length - span.Length - headerStart;
        return true;
    }

    private static ReadOnlySpan<char> SkipLeadingWhitespaceAndComments(ReadOnlySpan<char> span)
    {
        while (true)
        {
            var trimmed = span.TrimStart();

            if (trimmed.StartsWith("--", StringComparison.Ordinal))
            {
                var newlineIndex = trimmed.IndexOf('\n');
                trimmed = newlineIndex < 0 ? ReadOnlySpan<char>.Empty : trimmed[(newlineIndex + 1)..];
            }
            else if (trimmed.StartsWith("/*", StringComparison.Ordinal))
            {
                var endIndex = trimmed.IndexOf("*/", StringComparison.Ordinal);
                trimmed = endIndex < 0 ? ReadOnlySpan<char>.Empty : trimmed[(endIndex + 2)..];
            }
            else
            {
                return trimmed;
            }

            span = trimmed;
        }
    }

    private static bool TryConsumeKeyword(ref ReadOnlySpan<char> span, string keyword)
    {
        var trimmed = span.TrimStart();
        if (trimmed.Length < keyword.Length)
        {
            return false;
        }

        if (!trimmed[..keyword.Length].Equals(keyword, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (trimmed.Length > keyword.Length && !char.IsWhiteSpace(trimmed[keyword.Length]))
        {
            return false;
        }

        span = trimmed[keyword.Length..];
        return true;
    }

    public static IReadOnlyList<string> ExtractParameterNames(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql))
        {
            return Array.Empty<string>();
        }

        if (IsCreateOrAlterRoutineStatement(sql))
        {
            return Array.Empty<string>();
        }

        var names = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Variables the batch DECLAREs itself aren't query parameters - prompting for them and
        // then binding a same-named parameter makes SQL Server reject the DECLARE ("variable
        // name has already been declared"). A DECLARE statement declares its first @name plus
        // every @name after a top-level comma, until a ';' or the next statement keyword.
        var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var insideDeclare = false;
        var expectDeclaredName = false;
        var parenDepth = 0;

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
                expectDeclaredName = false;
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

            if (char.IsWhiteSpace(current))
            {
                continue;
            }

            if (char.IsLetterOrDigit(current) || current == '_')
            {
                var wordEnd = i + 1;
                while (wordEnd < sql.Length && (sql[wordEnd] == '_' || char.IsLetterOrDigit(sql[wordEnd])))
                {
                    wordEnd++;
                }

                var word = sql.AsSpan(i, wordEnd - i);
                if (word.Equals("DECLARE", StringComparison.OrdinalIgnoreCase))
                {
                    insideDeclare = true;
                    expectDeclaredName = true;
                    parenDepth = 0;
                }
                else
                {
                    expectDeclaredName = false;
                    if (insideDeclare && parenDepth == 0 && StatementKeywords.Contains(word.ToString()))
                    {
                        insideDeclare = false;
                    }
                }

                i = wordEnd - 1;
                continue;
            }

            if (current != '@')
            {
                expectDeclaredName = false;
                if (!insideDeclare)
                {
                    continue;
                }

                switch (current)
                {
                    case '(':
                        parenDepth++;
                        break;
                    case ')':
                        parenDepth = Math.Max(0, parenDepth - 1);
                        break;
                    case ',' when parenDepth == 0:
                        expectDeclaredName = true;
                        break;
                    case ';':
                        insideDeclare = false;
                        break;
                }

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
            if (expectDeclaredName)
            {
                declared.Add(parameterName);
                expectDeclaredName = false;
            }
            else if (seen.Add(parameterName))
            {
                names.Add(parameterName);
            }

            i = end - 1;
        }

        return names.Where(name => !declared.Contains(name)).ToList();
    }

    private static readonly HashSet<string> StatementKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "SELECT", "INSERT", "UPDATE", "DELETE", "MERGE", "EXEC", "EXECUTE", "SET", "IF", "ELSE",
        "WHILE", "BEGIN", "END", "RETURN", "PRINT", "USE", "TRUNCATE", "DROP", "CREATE", "ALTER",
        "OPEN", "FETCH", "CLOSE", "DEALLOCATE", "RAISERROR", "THROW", "WAITFOR", "GOTO", "BREAK", "CONTINUE"
    };

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
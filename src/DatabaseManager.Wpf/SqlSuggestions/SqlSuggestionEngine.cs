using System.Text.RegularExpressions;
using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.SqlSuggestions;

public sealed class SqlSuggestionEngine : ISqlSuggestionEngine
{
    private static readonly string[] DefaultSqlKeywordSuggestions =
    {
        "SELECT", "FROM", "WHERE", "GROUP BY", "ORDER BY", "HAVING", "JOIN", "INNER JOIN", "LEFT JOIN",
        "RIGHT JOIN", "FULL JOIN", "CROSS JOIN", "ON", "TOP", "DISTINCT", "INSERT INTO", "VALUES",
        "UPDATE", "SET", "DELETE", "TRUNCATE TABLE", "CREATE TABLE", "ALTER TABLE", "DROP TABLE",
        "CREATE PROCEDURE", "ALTER PROCEDURE", "DROP PROCEDURE", "EXEC", "BEGIN", "END", "CASE", "WHEN",
        "THEN", "ELSE", "DECLARE", "IF", "EXISTS", "IN", "BETWEEN", "LIKE", "UNION", "UNION ALL"
    };

    private static readonly string[] DefaultSqlFunctionSuggestions =
    {
        "COUNT()", "SUM()", "AVG()", "MIN()", "MAX()", "COALESCE()", "ISNULL()", "CAST()", "CONVERT()",
        "ROW_NUMBER() OVER (...)", "GETDATE()", "DATEDIFF()", "DATEADD()", "LEN()", "SUBSTRING()", "UPPER()", "LOWER()"
    };

    private static readonly string[] DefaultSqlSnippetSuggestions =
    {
        "SELECT TOP (100) * FROM [dbo].[TableName];",
        "INSERT INTO [dbo].[TableName] ([Column1]) VALUES (@Column1);",
        "UPDATE [dbo].[TableName] SET [Column1] = @Column1 WHERE [KeyColumn] = @KeyColumn;",
        "DELETE FROM [dbo].[TableName] WHERE [KeyColumn] = @KeyColumn;",
        "EXEC [dbo].[ProcedureName] @Param1 = @Param1;"
    };

    private static readonly Regex AliasRegex = new(
        @"\b(?:FROM|JOIN)\s+(?<obj>\[[^\]]+\]\.\[[^\]]+\]|\w+(?:\.\w+)?)\s+(?:AS\s+)?(?<alias>\w+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReferencedTableRegex = new(
        @"\b(?:FROM|JOIN)\s+(?<obj>\[[^\]]+\]\.\[[^\]]+\]|\w+(?:\.\w+)?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public IReadOnlyList<string> GetSuggestions(SqlSuggestionRequest request)
    {
        var normalizedToken = request.Token.Trim()
            .TrimStart('[')
            .TrimEnd(']');
        var contextKeyword = GetSuggestionContextKeyword(request.SqlText, request.TokenStart);
        var (qualifier, tokenSuffix) = SplitQualifier(normalizedToken);

        var aliasMap = BuildAliasMap(request.SqlText, request.TokenStart);

        var candidates = new Dictionary<string, SuggestionMetadata>(StringComparer.OrdinalIgnoreCase);

        void AddCandidate(string value, string category, string? schemaName = null)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (!candidates.ContainsKey(value))
            {
                candidates[value] = new SuggestionMetadata
                {
                    Category = category,
                    SchemaName = schemaName
                };
            }
        }

        var keywordSuggestions = request.Keywords.Count > 0 ? request.Keywords : DefaultSqlKeywordSuggestions;
        var functionSuggestions = request.Functions.Count > 0 ? request.Functions : DefaultSqlFunctionSuggestions;
        var snippetSuggestions = request.Snippets.Count > 0 ? request.Snippets : DefaultSqlSnippetSuggestions;

        foreach (var suggestion in keywordSuggestions)
        {
            AddCandidate(suggestion, "keyword");
        }

        foreach (var suggestion in functionSuggestions)
        {
            AddCandidate(suggestion, "function");
        }

        foreach (var suggestion in snippetSuggestions)
        {
            AddCandidate(suggestion, "snippet");
        }

        foreach (var table in request.Tables)
        {
            AddCandidate(table.FullName, "table", table.SchemaName);
            AddCandidate(table.TableName, "table", table.SchemaName);
            AddCandidate(table.SchemaName + ".", "schema", table.SchemaName);
        }

        foreach (var procedure in request.StoredProcedures)
        {
            AddCandidate(procedure.FullName, "procedure", procedure.SchemaName);
            AddCandidate(procedure.ProcedureName, "procedure", procedure.SchemaName);
        }

        foreach (var parameter in request.ProcedureParameters)
        {
            AddCandidate(parameter.ParameterName, "parameter");
        }

        foreach (var column in request.SelectedColumns)
        {
            AddCandidate(column.ColumnName, "column", column.SchemaName);
        }

        foreach (var fragment in request.RecentFragments)
        {
            AddCandidate(fragment, "recent");
        }

        foreach (var joinHint in BuildJoinHintSuggestions(request, contextKeyword, aliasMap))
        {
            AddCandidate(joinHint, "join-hint");
        }

        if (!string.IsNullOrWhiteSpace(qualifier))
        {
            var qualifierUpper = qualifier.ToUpperInvariant();

            if (aliasMap.TryGetValue(qualifierUpper, out var aliasTarget))
            {
                foreach (var column in request.SelectedColumns.Where(c =>
                             c.SchemaName.Equals(aliasTarget.SchemaName, StringComparison.OrdinalIgnoreCase)
                             && c.TableName.Equals(aliasTarget.TableName, StringComparison.OrdinalIgnoreCase)))
                {
                    AddCandidate(column.ColumnName, "column", column.SchemaName);
                }
            }
            else
            {
                foreach (var table in request.Tables.Where(t =>
                             t.SchemaName.Equals(qualifier, StringComparison.OrdinalIgnoreCase)))
                {
                    AddCandidate(table.TableName, "table", table.SchemaName);
                    AddCandidate(table.FullName, "table", table.SchemaName);
                }

                foreach (var procedure in request.StoredProcedures.Where(p =>
                             p.SchemaName.Equals(qualifier, StringComparison.OrdinalIgnoreCase)))
                {
                    AddCandidate(procedure.ProcedureName, "procedure", procedure.SchemaName);
                    AddCandidate(procedure.FullName, "procedure", procedure.SchemaName);
                }
            }
        }

        var lookupToken = string.IsNullOrWhiteSpace(tokenSuffix) ? normalizedToken : tokenSuffix;

        var results = candidates
            .Select(x => new
            {
                Value = x.Key,
                Rank = GetSuggestionRank(
                    x.Key,
                    x.Value,
                    lookupToken,
                    qualifier,
                    contextKeyword,
                    aliasMap)
            })
            .Where(x => x.Rank < 100)
            .OrderBy(x => x.Rank)
            .ThenBy(x => x.Value, StringComparer.OrdinalIgnoreCase)
            .Take(request.MaxResults)
            .Select(x => x.Value)
            .ToList();

        return results;
    }

    private static Dictionary<string, (string SchemaName, string TableName)> BuildAliasMap(string sqlText, int tokenStart)
    {
        var prefix = sqlText[..Math.Clamp(tokenStart, 0, sqlText.Length)];
        var map = new Dictionary<string, (string SchemaName, string TableName)>(StringComparer.OrdinalIgnoreCase);

        foreach (Match match in AliasRegex.Matches(prefix))
        {
            var alias = match.Groups["alias"].Value;
            var obj = match.Groups["obj"].Value;
            if (string.IsNullOrWhiteSpace(alias) || string.IsNullOrWhiteSpace(obj))
            {
                continue;
            }

            var (schema, table) = ParseObjectName(obj);
            if (string.IsNullOrWhiteSpace(schema) || string.IsNullOrWhiteSpace(table))
            {
                continue;
            }

            map[alias.ToUpperInvariant()] = (schema, table);
        }

        return map;
    }

    private static (string SchemaName, string TableName) ParseObjectName(string objectName)
    {
        var cleaned = objectName.Replace("[", string.Empty, StringComparison.Ordinal)
            .Replace("]", string.Empty, StringComparison.Ordinal);

        var parts = cleaned.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 1)
        {
            return ("dbo", parts[0]);
        }

        return (parts[0], parts[1]);
    }

    private static int GetSuggestionRank(
        string suggestion,
        SuggestionMetadata metadata,
        string token,
        string qualifier,
        string contextKeyword,
        IReadOnlyDictionary<string, (string SchemaName, string TableName)> aliasMap)
    {
        var rank = 0;

        if (contextKeyword is "FROM" or "JOIN" or "UPDATE" or "INTO" or "TABLE")
        {
            rank += metadata.Category is "table" or "schema" ? 0 : 20;
        }
        else if (contextKeyword is "EXEC" or "EXECUTE")
        {
            rank += metadata.Category is "procedure" or "parameter" ? 0 : 20;
        }
        else if (contextKeyword is "WHERE" or "ON")
        {
            rank += metadata.Category is "column" or "function" ? 0 : 20;
        }

        if (metadata.Category == "join-hint" && contextKeyword is ("JOIN" or "ON"))
        {
            rank -= 8;
        }

        if (metadata.Category == "recent")
        {
            rank += 6;
        }

        if (!string.IsNullOrWhiteSpace(qualifier))
        {
            if (aliasMap.ContainsKey(qualifier.ToUpperInvariant()))
            {
                rank += metadata.Category == "column" ? 0 : 25;
            }
            else if (metadata.SchemaName is not null && metadata.SchemaName.Equals(qualifier, StringComparison.OrdinalIgnoreCase))
            {
                rank -= 5;
            }
            else
            {
                rank += 10;
            }
        }

        if (string.IsNullOrWhiteSpace(token))
        {
            return rank + 10;
        }

        if (suggestion.StartsWith(token, StringComparison.OrdinalIgnoreCase))
        {
            return rank;
        }

        if (suggestion.Contains(token, StringComparison.OrdinalIgnoreCase))
        {
            return rank + 10;
        }

        return 100;
    }

    private static IReadOnlyList<string> BuildJoinHintSuggestions(
        SqlSuggestionRequest request,
        string contextKeyword,
        IReadOnlyDictionary<string, (string SchemaName, string TableName)> aliasMap)
    {
        if (contextKeyword is not ("JOIN" or "ON") || request.ForeignKeys.Count == 0)
        {
            return Array.Empty<string>();
        }

        var referencedTables = ExtractReferencedTables(request.SqlText, request.TokenStart);
        if (referencedTables.Count == 0)
        {
            return Array.Empty<string>();
        }

        var reverseAliasMap = BuildReverseAliasMap(aliasMap);
        var hints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var foreignKey in request.ForeignKeys)
        {
            var parentKey = ToTableKey(foreignKey.ParentSchemaName, foreignKey.ParentTableName);
            var referencedKey = ToTableKey(foreignKey.ReferencedSchemaName, foreignKey.ReferencedTableName);

            if (referencedTables.Contains(parentKey) && !referencedTables.Contains(referencedKey))
            {
                var snippet = BuildJoinSnippet(
                    joinSchema: foreignKey.ReferencedSchemaName,
                    joinTable: foreignKey.ReferencedTableName,
                    leftSchema: foreignKey.ParentSchemaName,
                    leftTable: foreignKey.ParentTableName,
                    leftColumn: foreignKey.ParentColumnName,
                    rightSchema: foreignKey.ReferencedSchemaName,
                    rightTable: foreignKey.ReferencedTableName,
                    rightColumn: foreignKey.ReferencedColumnName,
                    reverseAliasMap);

                hints.Add(snippet);
                continue;
            }

            if (referencedTables.Contains(referencedKey) && !referencedTables.Contains(parentKey))
            {
                var snippet = BuildJoinSnippet(
                    joinSchema: foreignKey.ParentSchemaName,
                    joinTable: foreignKey.ParentTableName,
                    leftSchema: foreignKey.ParentSchemaName,
                    leftTable: foreignKey.ParentTableName,
                    leftColumn: foreignKey.ParentColumnName,
                    rightSchema: foreignKey.ReferencedSchemaName,
                    rightTable: foreignKey.ReferencedTableName,
                    rightColumn: foreignKey.ReferencedColumnName,
                    reverseAliasMap);

                hints.Add(snippet);
            }
        }

        return hints.Take(12).ToList();
    }

    private static HashSet<string> ExtractReferencedTables(string sqlText, int tokenStart)
    {
        var prefix = sqlText[..Math.Clamp(tokenStart, 0, sqlText.Length)];
        var tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (Match match in ReferencedTableRegex.Matches(prefix))
        {
            var obj = match.Groups["obj"].Value;
            if (string.IsNullOrWhiteSpace(obj))
            {
                continue;
            }

            var (schema, table) = ParseObjectName(obj);
            tables.Add(ToTableKey(schema, table));
        }

        return tables;
    }

    private static Dictionary<string, string> BuildReverseAliasMap(
        IReadOnlyDictionary<string, (string SchemaName, string TableName)> aliasMap)
    {
        var reverseMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in aliasMap)
        {
            var key = ToTableKey(entry.Value.SchemaName, entry.Value.TableName);
            reverseMap[key] = entry.Key;
        }

        return reverseMap;
    }

    private static string BuildJoinSnippet(
        string joinSchema,
        string joinTable,
        string leftSchema,
        string leftTable,
        string leftColumn,
        string rightSchema,
        string rightTable,
        string rightColumn,
        IReadOnlyDictionary<string, string> reverseAliasMap)
    {
        var joinTarget = $"[{joinSchema}].[{joinTable}]";
        var leftQualifier = GetJoinQualifier(leftSchema, leftTable, reverseAliasMap);
        var rightQualifier = GetJoinQualifier(rightSchema, rightTable, reverseAliasMap);

        return $"JOIN {joinTarget} ON {leftQualifier}.[{leftColumn}] = {rightQualifier}.[{rightColumn}]";
    }

    private static string GetJoinQualifier(string schemaName, string tableName, IReadOnlyDictionary<string, string> reverseAliasMap)
    {
        var key = ToTableKey(schemaName, tableName);
        if (reverseAliasMap.TryGetValue(key, out var alias) && !string.IsNullOrWhiteSpace(alias))
        {
            return alias;
        }

        return $"[{schemaName}].[{tableName}]";
    }

    private static string ToTableKey(string schemaName, string tableName)
    {
        return $"{schemaName}.{tableName}";
    }

    private static string GetSuggestionContextKeyword(string sqlText, int tokenStart)
    {
        if (string.IsNullOrWhiteSpace(sqlText) || tokenStart <= 0)
        {
            return string.Empty;
        }

        var prefix = sqlText[..Math.Clamp(tokenStart, 0, sqlText.Length)];
        var match = Regex.Match(prefix, @"([A-Za-z_]+)\s*$", RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return string.Empty;
        }

        return match.Groups[1].Value.ToUpperInvariant();
    }

    private static (string Qualifier, string Suffix) SplitQualifier(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return (string.Empty, string.Empty);
        }

        var dotIndex = token.LastIndexOf('.');
        if (dotIndex < 0)
        {
            return (string.Empty, token);
        }

        var qualifier = token[..dotIndex].Trim('[', ']', ' ');
        var suffix = token[(dotIndex + 1)..];
        return (qualifier, suffix);
    }

    private sealed class SuggestionMetadata
    {
        public required string Category { get; init; }

        public string? SchemaName { get; init; }
    }
}

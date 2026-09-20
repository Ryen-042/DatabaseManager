using System.Globalization;
using System.Text;

namespace DatabaseManager.Core.Services;

/// <summary>
/// Pure SQL-text construction for RowEditService - no connection, no I/O, fully unit-testable.
/// Extracted so the WHERE-predicate building (including the null-safe match clauses used by
/// both by-primary-key updates and the no-PK delete safeguard) can be verified without a live
/// database, which RowEditService itself can't be since every one of its methods opens a
/// SqlConnection.
/// </summary>
public static class RowEditSqlBuilder
{
    public static string EscapeIdentifier(string value) => value.Replace("]", "]]", StringComparison.Ordinal);

    public static string BuildUpdateStatement(
        string schemaName,
        string tableName,
        IReadOnlyList<string> setColumns,
        IReadOnlyList<string> matchColumns,
        bool hasPrimaryKey)
    {
        var setClause = string.Join(", ", setColumns.Select(c => $"[{EscapeIdentifier(c)}] = @set_{c}"));
        var whereClause = BuildNullSafeWhereClause(matchColumns);
        var sql = $"UPDATE [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] SET {setClause} WHERE {whereClause};";

        if (!hasPrimaryKey)
        {
            sql += Environment.NewLine
                + "IF @@ROWCOUNT <> 1" + Environment.NewLine
                + "    THROW 50000, 'Cannot save changes because the row is not uniquely identifiable without a primary key.', 1;";
        }

        return sql;
    }

    public static string BuildInsertStatement(string schemaName, string tableName, IReadOnlyList<string> valueColumns)
    {
        var columnClause = string.Join(", ", valueColumns.Select(c => $"[{EscapeIdentifier(c)}]"));
        var valuesClause = string.Join(", ", valueColumns.Select(c => $"@ins_{c}"));
        return $"INSERT INTO [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] ({columnClause}) VALUES ({valuesClause});";
    }

    public static string BuildDeleteByPrimaryKeyStatement(string schemaName, string tableName, IReadOnlyList<string> primaryKeys)
    {
        var whereClause = string.Join(" AND ", primaryKeys.Select(c => $"[{EscapeIdentifier(c)}] = @key_{c}"));
        return $"DELETE FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] WHERE {whereClause};";
    }

    public static string BuildNullSafeWhereClause(IReadOnlyList<string> columns) => string.Join(" AND ", columns.Select(column =>
        $"((@key_{column} IS NULL AND [{EscapeIdentifier(column)}] IS NULL) OR [{EscapeIdentifier(column)}] = @key_{column})"));

    public static string BuildWhereClauseForSelectedColumns(IReadOnlyList<string> selectedColumns, string parameterPrefix = "") => string.Join(" AND ", selectedColumns.Select(column =>
        $"((@{parameterPrefix}{column} IS NULL AND [{EscapeIdentifier(column)}] IS NULL) OR [{EscapeIdentifier(column)}] = @{parameterPrefix}{column})"));

    public static string BuildCountSql(string schemaName, string tableName, IReadOnlyList<string> selectedColumns)
    {
        var whereClause = BuildWhereClauseForSelectedColumns(selectedColumns);
        return $"SELECT COUNT(1) FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] WHERE {whereClause};";
    }

    public static string BuildDeleteTopOneSql(string schemaName, string tableName, IReadOnlyList<string> selectedColumns)
    {
        var whereClause = BuildWhereClauseForSelectedColumns(selectedColumns);
        return $"DELETE TOP (1) FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}] WHERE {whereClause};";
    }

    /// <summary>
    /// Builds a human-reviewable DELETE script for the no-PK mismatch warning dialog - not
    /// executed directly, just copied to the clipboard when intended/matched row counts disagree.
    /// </summary>
    public static string BuildGeneratedDeleteSql(
        string schemaName,
        string tableName,
        IReadOnlyList<string> selectedColumns,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> selectedRows)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"-- Generated delete script for [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}]");
        sb.AppendLine("-- Review before execution.");
        sb.AppendLine();

        for (var i = 0; i < selectedRows.Count; i++)
        {
            var row = selectedRows[i];
            var predicates = new List<string>();

            foreach (var column in selectedColumns)
            {
                if (!row.TryGetValue(column, out var value))
                {
                    throw new InvalidOperationException($"Selected row does not include value for column '{column}'.");
                }

                predicates.Add(value is null or DBNull
                    ? $"[{EscapeIdentifier(column)}] IS NULL"
                    : $"[{EscapeIdentifier(column)}] = {ToSqlLiteral(value)}");
            }

            sb.AppendLine($"-- Row {i + 1}");
            sb.AppendLine($"DELETE TOP (1) FROM [{EscapeIdentifier(schemaName)}].[{EscapeIdentifier(tableName)}]");
            sb.AppendLine($"WHERE {string.Join(" AND ", predicates)};");
            sb.AppendLine("GO");
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public static string ToSqlLiteral(object value)
    {
        return value switch
        {
            string s => $"N'{s.Replace("'", "''", StringComparison.Ordinal)}'",
            char c => $"N'{c.ToString().Replace("'", "''", StringComparison.Ordinal)}'",
            bool b => b ? "1" : "0",
            DateTime dt => $"'{dt:yyyy-MM-dd HH:mm:ss.fffffff}'",
            DateTimeOffset dto => $"'{dto:yyyy-MM-dd HH:mm:ss.fffffff zzz}'",
            byte[] bytes => $"0x{Convert.ToHexString(bytes)}",
            Guid guid => $"'{guid:D}'",
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture) ?? "NULL",
            _ => $"N'{value.ToString()?.Replace("'", "''", StringComparison.Ordinal)}'"
        };
    }
}

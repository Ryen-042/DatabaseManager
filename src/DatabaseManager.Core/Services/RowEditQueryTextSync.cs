using System.Text.RegularExpressions;

namespace DatabaseManager.Core.Services;

/// <summary>
/// Pure structured-inputs to/from SQL-text conversion for the Edit Rows tab's two synchronized
/// modes (Top/Filter/Order By inputs vs. a directly-editable SQL box). No UI dependency, so it
/// doesn't itself implement "custom mode wins, don't overwrite" - callers gate BuildQuery calls
/// on their own custom-mode flag exactly like MainWindow's _isEditRowsCustomQueryMode does today;
/// that's what actually preserves any comments/formatting the user has in the custom text, since
/// this class is never asked to regenerate over them while custom mode is active.
/// </summary>
public static class RowEditQueryTextSync
{
    private static readonly Regex EditRowsQueryRegex = new(
        @"^\s*SELECT\s+(?:TOP\s*\(\s*(?<top>\d+)\s*\)\s+)?\*\s+FROM\s+(?<from>\[[^\]]+\]\.\[[^\]]+\]|\S+)(?:\s+WHERE\s+(?<where>.*?))?(?:\s+ORDER\s+BY\s+(?<order>.*?))?\s*;?\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public static string BuildQuery(string schemaName, string tableName, int topRows, string? filter, string? orderBy)
    {
        var sql = $"SELECT TOP ({topRows}) *{Environment.NewLine}FROM [{schemaName}].[{tableName}]";

        if (!string.IsNullOrWhiteSpace(filter))
        {
            sql += $"{Environment.NewLine}WHERE {filter}";
        }

        if (!string.IsNullOrWhiteSpace(orderBy))
        {
            sql += $"{Environment.NewLine}ORDER BY {orderBy}";
        }

        sql += ";";
        return sql;
    }

    public static bool TryParse(string sql, out int topRows, out string? filter, out string? orderBy)
    {
        topRows = 200;
        filter = null;
        orderBy = null;

        var match = EditRowsQueryRegex.Match(sql ?? string.Empty);
        if (!match.Success)
        {
            return false;
        }

        if (match.Groups["top"].Success
            && (!int.TryParse(match.Groups["top"].Value, out topRows) || topRows <= 0))
        {
            return false;
        }

        filter = match.Groups["where"].Success
            ? match.Groups["where"].Value.Trim()
            : null;

        orderBy = match.Groups["order"].Success
            ? match.Groups["order"].Value.Trim().TrimEnd(';')
            : null;

        if (string.IsNullOrWhiteSpace(filter))
        {
            filter = null;
        }

        if (string.IsNullOrWhiteSpace(orderBy))
        {
            orderBy = null;
        }

        return true;
    }
}

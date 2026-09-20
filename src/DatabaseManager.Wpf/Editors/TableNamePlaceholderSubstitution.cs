using System.Text.RegularExpressions;

namespace DatabaseManager.Wpf.Editors;

/// <summary>
/// Replaces the "TableName" placeholder (bare, bracketed, or schema-qualified like
/// "[dbo].[TableName]" - the exact shape the built-in SQL suggestion snippets use) with the
/// selected table's real qualified name. Matching the schema-qualified form as a single unit
/// matters: replacing only the "[TableName]" portion of "[dbo].[TableName]" would leave the
/// snippet's hardcoded "dbo" behind, producing an invalid three-part name.
/// </summary>
public static class TableNamePlaceholderSubstitution
{
    private static readonly Regex PlaceholderRegex = new(
        @"(?:(?:\[\w+\]|\w+)\.)?\[TableName\]|(?:(?:\[\w+\]|\w+)\.)?TableName",
        RegexOptions.Compiled);

    /// <summary>
    /// Returns the substituted text, or null if <paramref name="text"/> contains no placeholder
    /// to replace (so callers can skip a no-op text/caret update).
    /// </summary>
    public static string? TrySubstitute(string text, string schemaName, string tableName)
    {
        if (!text.Contains("TableName", StringComparison.Ordinal))
        {
            return null;
        }

        var newText = PlaceholderRegex.Replace(text, $"[{schemaName}].[{tableName}]");
        return string.Equals(text, newText, StringComparison.Ordinal) ? null : newText;
    }
}

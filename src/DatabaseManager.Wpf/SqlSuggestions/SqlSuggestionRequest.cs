using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.SqlSuggestions;

public sealed class SqlSuggestionRequest
{
    public required string SqlText { get; init; }

    public required string Token { get; init; }

    public required int TokenStart { get; init; }

    public IReadOnlyList<string> Keywords { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Functions { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Snippets { get; init; } = Array.Empty<string>();

    public required IReadOnlyList<TableSchemaInfo> Tables { get; init; }

    public required IReadOnlyList<ColumnSchemaInfo> SelectedColumns { get; init; }

    public required IReadOnlyList<StoredProcedureSchemaInfo> StoredProcedures { get; init; }

    public required IReadOnlyList<StoredProcedureParameterInfo> ProcedureParameters { get; init; }

    public IReadOnlyList<ForeignKeySchemaInfo> ForeignKeys { get; init; } = Array.Empty<ForeignKeySchemaInfo>();

    public IReadOnlyList<string> RecentFragments { get; init; } = Array.Empty<string>();

    public int MaxResults { get; init; } = 30;
}

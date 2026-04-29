using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.SqlSuggestions;

public sealed class SqlCompletionCatalogSnapshot
{
    public required IReadOnlyList<string> Keywords { get; init; }

    public required IReadOnlyList<string> Functions { get; init; }

    public required IReadOnlyList<string> Snippets { get; init; }

    public required IReadOnlyList<TableSchemaInfo> Tables { get; init; }

    public required IReadOnlyList<StoredProcedureSchemaInfo> StoredProcedures { get; init; }

    public required IReadOnlyList<ColumnSchemaInfo> GlobalColumns { get; init; }

    public required IReadOnlyList<StoredProcedureParameterInfo> ProcedureParameters { get; init; }

    public required IReadOnlyList<ForeignKeySchemaInfo> ForeignKeys { get; init; }

    public DateTimeOffset LastRefreshUtc { get; init; }
}

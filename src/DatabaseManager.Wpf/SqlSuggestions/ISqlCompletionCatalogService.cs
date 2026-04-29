using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.SqlSuggestions;

public interface ISqlCompletionCatalogService
{
    SqlCompletionCatalogSnapshot GetSnapshot();

    void RefreshSchemaMetadata(
        IReadOnlyList<TableSchemaInfo> tables,
        IReadOnlyList<StoredProcedureSchemaInfo> storedProcedures,
        IReadOnlyList<ForeignKeySchemaInfo> foreignKeys);

    void RefreshTableColumns(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns);

    void RefreshProcedureParameters(StoredProcedureSchemaInfo procedure, IReadOnlyList<StoredProcedureParameterInfo> parameters);
}

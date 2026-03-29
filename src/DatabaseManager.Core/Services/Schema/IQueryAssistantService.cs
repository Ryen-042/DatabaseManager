using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Core.Services.Schema;

public interface IQueryAssistantService
{
    string BuildSelectTopQuery(TableSchemaInfo table, int topRows = 100);

    string BuildInsertQuery(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns);

    string BuildUpdateQuery(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns);

    string BuildDeleteQuery(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns);

    string BuildExecuteProcedureQuery(StoredProcedureSchemaInfo procedure, IReadOnlyList<StoredProcedureParameterInfo> parameters);

    string BuildTableSchemaText(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns);
}

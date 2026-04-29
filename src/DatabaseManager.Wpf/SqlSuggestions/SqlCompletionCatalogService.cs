using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Wpf.Editors;

namespace DatabaseManager.Wpf.SqlSuggestions;

public sealed class SqlCompletionCatalogService : ISqlCompletionCatalogService
{
    private static readonly string[] SqlFunctionSuggestions =
    {
        "COUNT()", "SUM()", "AVG()", "MIN()", "MAX()", "COALESCE()", "ISNULL()", "CAST()", "CONVERT()",
        "ROW_NUMBER() OVER (...)", "GETDATE()", "DATEDIFF()", "DATEADD()", "LEN()", "SUBSTRING()", "UPPER()", "LOWER()"
    };

    private static readonly string[] SqlSnippetSuggestions =
    {
        "SELECT TOP (100) * FROM [dbo].[TableName];",
        "INSERT INTO [dbo].[TableName] ([Column1]) VALUES (@Column1);",
        "UPDATE [dbo].[TableName] SET [Column1] = @Column1 WHERE [KeyColumn] = @KeyColumn;",
        "DELETE FROM [dbo].[TableName] WHERE [KeyColumn] = @KeyColumn;",
        "EXEC [dbo].[ProcedureName] @Param1 = @Param1;"
    };

    private readonly object _sync = new();
    private List<TableSchemaInfo> _tables = new();
    private List<StoredProcedureSchemaInfo> _storedProcedures = new();
    private List<ForeignKeySchemaInfo> _foreignKeys = new();
    private readonly Dictionary<string, List<ColumnSchemaInfo>> _columnsByTable = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<StoredProcedureParameterInfo>> _parametersByProcedure = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<ColumnSchemaInfo> _globalColumns = new();
    private readonly List<StoredProcedureParameterInfo> _procedureParameters = new();
    private DateTimeOffset _lastRefreshUtc = DateTimeOffset.MinValue;

    public SqlCompletionCatalogSnapshot GetSnapshot()
    {
        lock (_sync)
        {
            return new SqlCompletionCatalogSnapshot
            {
                Keywords = SqlLanguageKeywords.Keywords,
                Functions = SqlFunctionSuggestions,
                Snippets = SqlSnippetSuggestions,
                Tables = _tables.ToList(),
                StoredProcedures = _storedProcedures.ToList(),
                GlobalColumns = _globalColumns.ToList(),
                ProcedureParameters = _procedureParameters.ToList(),
                ForeignKeys = _foreignKeys.ToList(),
                LastRefreshUtc = _lastRefreshUtc
            };
        }
    }

    public void RefreshSchemaMetadata(
        IReadOnlyList<TableSchemaInfo> tables,
        IReadOnlyList<StoredProcedureSchemaInfo> storedProcedures,
        IReadOnlyList<ForeignKeySchemaInfo> foreignKeys)
    {
        lock (_sync)
        {
            _tables = tables.ToList();
            _storedProcedures = storedProcedures.ToList();
            _foreignKeys = foreignKeys.ToList();

            // Remove stale per-table caches that no longer exist in the schema.
            var existingTableKeys = _tables
                .Select(ToTableKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var key in _columnsByTable.Keys.Where(x => !existingTableKeys.Contains(x)).ToList())
            {
                _columnsByTable.Remove(key);
            }

            var existingProcedureKeys = _storedProcedures
                .Select(ToProcedureKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var key in _parametersByProcedure.Keys.Where(x => !existingProcedureKeys.Contains(x)).ToList())
            {
                _parametersByProcedure.Remove(key);
            }

            RebuildDerivedCachesLocked();
            _lastRefreshUtc = DateTimeOffset.UtcNow;
        }
    }

    public void RefreshTableColumns(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var key = ToTableKey(table);
        lock (_sync)
        {
            _columnsByTable[key] = columns.ToList();
            RebuildDerivedCachesLocked();
            _lastRefreshUtc = DateTimeOffset.UtcNow;
        }
    }

    public void RefreshProcedureParameters(StoredProcedureSchemaInfo procedure, IReadOnlyList<StoredProcedureParameterInfo> parameters)
    {
        var key = ToProcedureKey(procedure);
        lock (_sync)
        {
            _parametersByProcedure[key] = parameters.ToList();
            RebuildDerivedCachesLocked();
            _lastRefreshUtc = DateTimeOffset.UtcNow;
        }
    }

    private void RebuildDerivedCachesLocked()
    {
        _globalColumns.Clear();
        _globalColumns.AddRange(_columnsByTable.Values
            .SelectMany(x => x)
            .GroupBy(x => $"{x.SchemaName}.{x.TableName}.{x.ColumnName}", StringComparer.OrdinalIgnoreCase)
            .Select(x => x.First()));

        _procedureParameters.Clear();
        _procedureParameters.AddRange(_parametersByProcedure.Values
            .SelectMany(x => x)
            .GroupBy(x => x.ParameterName, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.First()));
    }

    private static string ToTableKey(TableSchemaInfo table)
    {
        return $"{table.SchemaName}.{table.TableName}";
    }

    private static string ToProcedureKey(StoredProcedureSchemaInfo procedure)
    {
        return $"{procedure.SchemaName}.{procedure.ProcedureName}";
    }
}

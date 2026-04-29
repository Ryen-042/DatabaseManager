using System.Text;
using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Core.Services.Schema;

public sealed class SqlQueryAssistantService : IQueryAssistantService
{
    public string BuildSelectTopQuery(TableSchemaInfo table, int topRows = 100)
    {
        return $"SELECT TOP ({topRows}) *{Environment.NewLine}FROM {table.FullName};";
    }

    public string BuildInsertQuery(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var writableColumns = columns
            .Where(c => !c.IsIdentity)
            .OrderBy(c => c.OrdinalPosition)
            .ToList();

        var columnNames = string.Join($",{Environment.NewLine}    ", writableColumns.Select(c => $"[{c.ColumnName}]"));
        var values = string.Join($",{Environment.NewLine}    ", writableColumns.Select(c => $"@{c.ColumnName}"));

        return $"INSERT INTO {table.FullName} ({Environment.NewLine}    {columnNames}{Environment.NewLine}){Environment.NewLine}VALUES ({Environment.NewLine}    {values}{Environment.NewLine});";
    }

    public string BuildUpdateQuery(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var writableColumns = columns
            .Where(c => !c.IsIdentity && !c.IsPrimaryKey)
            .OrderBy(c => c.OrdinalPosition)
            .ToList();

        var primaryKeys = columns
            .Where(c => c.IsPrimaryKey)
            .OrderBy(c => c.OrdinalPosition)
            .ToList();

        var setClause = string.Join($",{Environment.NewLine}", writableColumns.Select(c => $"    [{c.ColumnName}] = @{c.ColumnName}"));
        var whereClause = BuildWhereClause(primaryKeys);

        return $"UPDATE {table.FullName}{Environment.NewLine}SET{Environment.NewLine}{setClause}{Environment.NewLine}WHERE {whereClause};";
    }

    public string BuildDeleteQuery(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var primaryKeys = columns
            .Where(c => c.IsPrimaryKey)
            .OrderBy(c => c.OrdinalPosition)
            .ToList();

        var whereClause = BuildWhereClause(primaryKeys);
        return $"DELETE FROM {table.FullName}{Environment.NewLine}WHERE {whereClause};";
    }

    public string BuildExecuteProcedureQuery(StoredProcedureSchemaInfo procedure, IReadOnlyList<StoredProcedureParameterInfo> parameters)
    {
        var ordered = parameters
            .Where(p => !p.IsReturnValue)
            .OrderBy(p => p.OrdinalPosition)
            .ToList();

        if (ordered.Count == 0)
        {
            return $"EXEC {procedure.FullName};";
        }

        var args = string.Join($",{Environment.NewLine}    ", ordered.Select(p =>
            p.IsOutput
                ? $"{p.ParameterName} = @{NormalizeParameterName(p.ParameterName)} OUTPUT"
                : $"{p.ParameterName} = @{NormalizeParameterName(p.ParameterName)}"));

        return $"EXEC {procedure.FullName}{Environment.NewLine}    {args};";
    }

    public string BuildTableSchemaText(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"-- Schema for {table.FullName}");
        sb.AppendLine($"CREATE TABLE {table.FullName}");
        sb.AppendLine("(");

        var ordered = columns.OrderBy(c => c.OrdinalPosition).ToList();
        for (var i = 0; i < ordered.Count; i++)
        {
            var c = ordered[i];
            var suffix = i == ordered.Count - 1 && !ordered.Any(x => x.IsPrimaryKey) ? string.Empty : ",";

            var dataType = GetTypeDeclaration(c);
            var identity = c.IsIdentity ? " IDENTITY(1,1)" : string.Empty;
            var nullable = c.IsNullable ? "NULL" : "NOT NULL";

            sb.AppendLine($"    [{c.ColumnName}] {dataType}{identity} {nullable}{suffix}");
        }

        var primaryKeys = ordered.Where(x => x.IsPrimaryKey).ToList();
        if (primaryKeys.Count > 0)
        {
            var keyColumns = string.Join(", ", primaryKeys.Select(x => $"[{x.ColumnName}]"));
            sb.AppendLine($"    CONSTRAINT [PK_{table.TableName}] PRIMARY KEY ({keyColumns})");
        }

        sb.AppendLine(");");
        return sb.ToString();
    }

    public string BuildDropTableScript(TableSchemaInfo table)
    {
        return $"IF OBJECT_ID(N'{table.FullName}', N'U') IS NOT NULL{Environment.NewLine}BEGIN{Environment.NewLine}    DROP TABLE {table.FullName};{Environment.NewLine}END;";
    }

    public string BuildDropAndCreateTableScript(TableSchemaInfo table, IReadOnlyList<ColumnSchemaInfo> columns)
    {
        var dropScript = BuildDropTableScript(table);
        var createScript = BuildTableSchemaText(table, columns);

        return $"{dropScript}{Environment.NewLine}{Environment.NewLine}GO{Environment.NewLine}{Environment.NewLine}{createScript}";
    }

    public string BuildDropProcedureScript(StoredProcedureSchemaInfo procedure)
    {
        return $"IF OBJECT_ID(N'{procedure.FullName}', N'P') IS NOT NULL{Environment.NewLine}BEGIN{Environment.NewLine}    DROP PROCEDURE {procedure.FullName};{Environment.NewLine}END;";
    }

    public string BuildAlterProcedureScript(StoredProcedureSchemaInfo procedure, string? currentDefinition)
    {
        if (string.IsNullOrWhiteSpace(currentDefinition))
        {
            return $"-- Procedure definition is unavailable for {procedure.FullName}.{Environment.NewLine}-- Ensure your login has VIEW DEFINITION permission or script manually.{Environment.NewLine}{Environment.NewLine}ALTER PROCEDURE {procedure.FullName}{Environment.NewLine}AS{Environment.NewLine}BEGIN{Environment.NewLine}    SET NOCOUNT ON;{Environment.NewLine}    -- TODO: Add procedure body.{Environment.NewLine}END;";
        }

        var definition = currentDefinition.Trim();
        var createToken = "CREATE PROCEDURE";

        var index = definition.IndexOf(createToken, StringComparison.OrdinalIgnoreCase);
        if (index >= 0)
        {
            return definition.Remove(index, createToken.Length).Insert(index, "ALTER PROCEDURE");
        }

        return $"-- Existing definition did not contain CREATE PROCEDURE token; review before execution.{Environment.NewLine}{definition}";
    }

    private static string BuildWhereClause(IReadOnlyList<ColumnSchemaInfo> primaryKeys)
    {
        if (primaryKeys.Count == 0)
        {
            return "/* TODO: Add safe WHERE clause */ 1 = 0";
        }

        return string.Join(" AND ", primaryKeys.Select(c => $"[{c.ColumnName}] = @{c.ColumnName}"));
    }

    private static string GetTypeDeclaration(ColumnSchemaInfo column)
    {
        var type = column.DataType.ToLowerInvariant();

        return type switch
        {
            "varchar" or "char" or "varbinary" or "binary" => $"{column.DataType}({(column.MaxLength == -1 ? "MAX" : column.MaxLength)})",
            "nvarchar" or "nchar" => $"{column.DataType}({(column.MaxLength == -1 ? "MAX" : column.MaxLength / 2)})",
            "decimal" or "numeric" => $"{column.DataType}({column.Precision},{column.Scale})",
            "datetime2" or "datetimeoffset" or "time" => $"{column.DataType}({column.Scale})",
            _ => column.DataType
        };
    }

    private static string NormalizeParameterName(string name)
    {
        return name.StartsWith("@", StringComparison.Ordinal) ? name[1..] : name;
    }
}

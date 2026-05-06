using DatabaseManager.Core.Models.Schema;
using Microsoft.Data.SqlClient;

namespace DatabaseManager.Core.Services.Schema;

public sealed class SqlServerSchemaService : IDatabaseSchemaService
{
    public async Task<IReadOnlyList<string>> GetDatabasesAsync(string connectionString, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT name
            FROM sys.databases
            WHERE state = 0
              AND HAS_DBACCESS(name) = 1
            ORDER BY name;
            """;

        var list = new List<string>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            if (!reader.IsDBNull(0))
            {
                list.Add(reader.GetString(0));
            }
        }

        return list;
    }

    public async Task<IReadOnlyList<TableSchemaInfo>> GetTablesAsync(string connectionString, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT
                s.name AS SchemaName,
                o.name AS TableName,
                CAST(SUM(CASE WHEN o.type = 'U' AND p.index_id IN (0,1) THEN p.rows ELSE 0 END) AS BIGINT) AS RowCount
            FROM sys.objects o
            INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
            LEFT JOIN sys.partitions p ON p.object_id = o.object_id
            WHERE o.is_ms_shipped = 0
              AND o.type IN ('U', 'V', 'SN')
            GROUP BY s.name, o.name
            ORDER BY s.name, o.name;
            """;

        const string fallbackSql = """
            SELECT
                TABLE_SCHEMA AS SchemaName,
                TABLE_NAME AS TableName,
                CAST(0 AS BIGINT) AS RowCount
            FROM INFORMATION_SCHEMA.TABLES
            ORDER BY TABLE_SCHEMA, TABLE_NAME;
            """;

        var list = new List<TableSchemaInfo>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        if (!await TryReadTablesAsync(sql, connection, list, cancellationToken) || list.Count == 0)
        {
            list.Clear();
            await TryReadTablesAsync(fallbackSql, connection, list, cancellationToken);
        }

        if (list.Count == 0)
        {
            ReadTablesFromAdoNetSchema(connection, list);
        }

        if (list.Count == 0)
        {
            await ReadTablesFromStoredProcedureAsync(connection, list, cancellationToken);
        }

        list = list
            .GroupBy(x => x.FullName, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(x => x.SchemaName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.TableName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return list;
    }

    public async Task<IReadOnlyList<ForeignKeySchemaInfo>> GetForeignKeysAsync(string connectionString, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT
                fk.name AS ConstraintName,
                ps.name AS ParentSchemaName,
                pt.name AS ParentTableName,
                pc.name AS ParentColumnName,
                rs.name AS ReferencedSchemaName,
                rt.name AS ReferencedTableName,
                rc.name AS ReferencedColumnName
            FROM sys.foreign_keys fk
            INNER JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id
            INNER JOIN sys.tables pt ON pt.object_id = fk.parent_object_id
            INNER JOIN sys.schemas ps ON ps.schema_id = pt.schema_id
            INNER JOIN sys.columns pc ON pc.object_id = pt.object_id AND pc.column_id = fkc.parent_column_id
            INNER JOIN sys.tables rt ON rt.object_id = fk.referenced_object_id
            INNER JOIN sys.schemas rs ON rs.schema_id = rt.schema_id
            INNER JOIN sys.columns rc ON rc.object_id = rt.object_id AND rc.column_id = fkc.referenced_column_id
            ORDER BY ps.name, pt.name, fk.name, fkc.constraint_column_id;
            """;

        var relationships = new List<ForeignKeySchemaInfo>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            relationships.Add(new ForeignKeySchemaInfo
            {
                ConstraintName = reader.GetString(0),
                ParentSchemaName = reader.GetString(1),
                ParentTableName = reader.GetString(2),
                ParentColumnName = reader.GetString(3),
                ReferencedSchemaName = reader.GetString(4),
                ReferencedTableName = reader.GetString(5),
                ReferencedColumnName = reader.GetString(6)
            });
        }

        return relationships;
    }

    public async Task<IReadOnlyList<ColumnSchemaInfo>> GetColumnsAsync(string connectionString, string schemaName, string tableName, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT
                s.name AS SchemaName,
                o.name AS TableName,
                c.name AS ColumnName,
                ty.name AS DataType,
                c.column_id AS OrdinalPosition,
                c.is_nullable AS IsNullable,
                c.is_identity AS IsIdentity,
                CAST(CASE WHEN pk.column_id IS NULL THEN 0 ELSE 1 END AS bit) AS IsPrimaryKey,
                c.max_length AS MaxLength,
                c.precision AS Precision,
                c.scale AS Scale
            FROM sys.columns c
            INNER JOIN sys.objects o ON o.object_id = c.object_id
            INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
            INNER JOIN sys.types ty ON ty.user_type_id = c.user_type_id
            LEFT JOIN (
                SELECT ic.object_id, ic.column_id
                FROM sys.indexes i
                INNER JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
                INNER JOIN sys.key_constraints kc ON kc.parent_object_id = i.object_id AND kc.unique_index_id = i.index_id
                WHERE kc.type = 'PK'
            ) pk ON pk.object_id = c.object_id AND pk.column_id = c.column_id
            WHERE s.name = @schemaName AND o.name = @tableName
              AND o.type IN ('U', 'V')
            ORDER BY c.column_id;
            """;

        var list = new List<ColumnSchemaInfo>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@schemaName", schemaName);
        command.Parameters.AddWithValue("@tableName", tableName);

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new ColumnSchemaInfo
            {
                SchemaName = reader.GetString(0),
                TableName = reader.GetString(1),
                ColumnName = reader.GetString(2),
                DataType = reader.GetString(3),
                OrdinalPosition = reader.GetInt32(4),
                IsNullable = reader.GetBoolean(5),
                IsIdentity = reader.GetBoolean(6),
                IsPrimaryKey = reader.GetBoolean(7),
                MaxLength = reader.GetInt16(8),
                Precision = reader.GetByte(9),
                Scale = reader.GetByte(10)
            });
        }

        if (list.Count > 0)
        {
            return list;
        }

        const string describeSql = """
            SELECT
                [name],
                system_type_name,
                column_ordinal,
                is_nullable
            FROM sys.dm_exec_describe_first_result_set(@query, NULL, 0)
            WHERE [name] IS NOT NULL
            ORDER BY column_ordinal;
            """;

        var safeSchema = schemaName.Replace("]", "]]", StringComparison.Ordinal);
        var safeObject = tableName.Replace("]", "]]", StringComparison.Ordinal);
        var queryForDescribe = $"SELECT TOP (0) * FROM [{safeSchema}].[{safeObject}]";

        await using var describeCommand = new SqlCommand(describeSql, connection);
        describeCommand.Parameters.AddWithValue("@query", queryForDescribe);

        await using var describeReader = await describeCommand.ExecuteReaderAsync(cancellationToken);

        while (await describeReader.ReadAsync(cancellationToken))
        {
            list.Add(new ColumnSchemaInfo
            {
                SchemaName = schemaName,
                TableName = tableName,
                ColumnName = describeReader.GetString(0),
                DataType = describeReader.IsDBNull(1) ? "sql_variant" : describeReader.GetString(1),
                OrdinalPosition = describeReader.GetInt32(2),
                IsNullable = describeReader.GetBoolean(3),
                IsIdentity = false,
                IsPrimaryKey = false,
                MaxLength = 0,
                Precision = 0,
                Scale = 0
            });
        }

        return list;
    }

    public async Task<IReadOnlyList<StoredProcedureSchemaInfo>> GetStoredProceduresAsync(string connectionString, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT s.name AS SchemaName, p.name AS ProcedureName
            FROM sys.procedures p
            INNER JOIN sys.schemas s ON s.schema_id = p.schema_id
            WHERE p.is_ms_shipped = 0
            ORDER BY s.name, p.name;
            """;

        const string fallbackSql = """
            SELECT
                SPECIFIC_SCHEMA AS SchemaName,
                SPECIFIC_NAME AS ProcedureName
            FROM INFORMATION_SCHEMA.ROUTINES
            WHERE ROUTINE_TYPE = 'PROCEDURE'
            ORDER BY SPECIFIC_SCHEMA, SPECIFIC_NAME;
            """;

        var list = new List<StoredProcedureSchemaInfo>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        if (!await TryReadProceduresAsync(sql, connection, list, cancellationToken) || list.Count == 0)
        {
            list.Clear();
            await TryReadProceduresAsync(fallbackSql, connection, list, cancellationToken);
        }

        return list;
    }

    public async Task<IReadOnlyList<StoredProcedureParameterInfo>> GetStoredProcedureParametersAsync(
        string connectionString,
        string schemaName,
        string procedureName,
        CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT
                p.name AS ParameterName,
                ty.name AS DataType,
                p.parameter_id AS OrdinalPosition,
                p.is_output AS IsOutput,
                p.has_default_value AS HasDefaultValue,
                p.max_length AS MaxLength,
                p.precision AS Precision,
                p.scale AS Scale
            FROM sys.parameters p
            INNER JOIN sys.procedures sp ON sp.object_id = p.object_id
            INNER JOIN sys.schemas s ON s.schema_id = sp.schema_id
            INNER JOIN sys.types ty ON ty.user_type_id = p.user_type_id
            WHERE s.name = @schemaName AND sp.name = @procedureName
            ORDER BY p.parameter_id;
            """;

        var list = new List<StoredProcedureParameterInfo>();

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@schemaName", schemaName);
        command.Parameters.AddWithValue("@procedureName", procedureName);

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new StoredProcedureParameterInfo
            {
                ParameterName = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                DataType = reader.GetString(1),
                OrdinalPosition = reader.GetInt32(2),
                IsOutput = reader.GetBoolean(3),
                HasDefaultValue = reader.GetBoolean(4),
                MaxLength = reader.GetInt16(5),
                Precision = reader.GetByte(6),
                Scale = reader.GetByte(7)
            });
        }

        return list;
    }

    public async Task<string?> GetStoredProcedureDefinitionAsync(
        string connectionString,
        string schemaName,
        string procedureName,
        CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT sm.definition
            FROM sys.procedures p
            INNER JOIN sys.schemas s ON s.schema_id = p.schema_id
            LEFT JOIN sys.sql_modules sm ON sm.object_id = p.object_id
            WHERE s.name = @schemaName AND p.name = @procedureName;
            """;

        const string fallbackSql = """
            SELECT OBJECT_DEFINITION(OBJECT_ID(QUOTENAME(@schemaName) + '.' + QUOTENAME(@procedureName)));
            """;

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancellationToken);

        var definition = await ExecuteDefinitionScalarAsync(sql, connection, schemaName, procedureName, cancellationToken);
        if (!string.IsNullOrWhiteSpace(definition))
        {
            return definition;
        }

        return await ExecuteDefinitionScalarAsync(fallbackSql, connection, schemaName, procedureName, cancellationToken);
    }

    private static async Task<bool> TryReadTablesAsync(
        string sql,
        SqlConnection connection,
        ICollection<TableSchemaInfo> results,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var command = new SqlCommand(sql, connection);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                results.Add(new TableSchemaInfo
                {
                    SchemaName = reader.GetString(0),
                    TableName = reader.GetString(1),
                    RowCount = reader.IsDBNull(2) ? 0 : reader.GetInt64(2)
                });
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static async Task<bool> TryReadProceduresAsync(
        string sql,
        SqlConnection connection,
        ICollection<StoredProcedureSchemaInfo> results,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var command = new SqlCommand(sql, connection);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                results.Add(new StoredProcedureSchemaInfo
                {
                    SchemaName = reader.GetString(0),
                    ProcedureName = reader.GetString(1)
                });
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void ReadTablesFromAdoNetSchema(SqlConnection connection, ICollection<TableSchemaInfo> results)
    {
        var schemaTable = connection.GetSchema("Tables");

        foreach (System.Data.DataRow row in schemaTable.Rows)
        {
            var schemaName = row["TABLE_SCHEMA"]?.ToString();
            var tableName = row["TABLE_NAME"]?.ToString();
            var tableType = row["TABLE_TYPE"]?.ToString();

            if (string.IsNullOrWhiteSpace(schemaName) || string.IsNullOrWhiteSpace(tableName))
            {
                continue;
            }

            if (!string.Equals(tableType, "BASE TABLE", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(tableType, "VIEW", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(tableType, "TABLE", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            results.Add(new TableSchemaInfo
            {
                SchemaName = schemaName,
                TableName = tableName,
                RowCount = 0
            });
        }
    }

    private static async Task ReadTablesFromStoredProcedureAsync(
        SqlConnection connection,
        ICollection<TableSchemaInfo> results,
        CancellationToken cancellationToken)
    {
        await using var command = new SqlCommand("sp_tables", connection)
        {
            CommandType = System.Data.CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@table_type", "'TABLE','VIEW'");

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            // Expected columns from sp_tables: TABLE_QUALIFIER, TABLE_OWNER, TABLE_NAME, TABLE_TYPE, REMARKS
            var schemaName = reader.IsDBNull(1) ? null : reader.GetString(1);
            var tableName = reader.IsDBNull(2) ? null : reader.GetString(2);
            var tableType = reader.IsDBNull(3) ? null : reader.GetString(3);

            if (string.IsNullOrWhiteSpace(schemaName) || string.IsNullOrWhiteSpace(tableName))
            {
                continue;
            }

            if (!string.Equals(tableType, "TABLE", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(tableType, "VIEW", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            results.Add(new TableSchemaInfo
            {
                SchemaName = schemaName,
                TableName = tableName,
                RowCount = 0
            });
        }
    }

    private static async Task<string?> ExecuteDefinitionScalarAsync(
        string sql,
        SqlConnection connection,
        string schemaName,
        string procedureName,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var command = new SqlCommand(sql, connection);
            command.Parameters.AddWithValue("@schemaName", schemaName);
            command.Parameters.AddWithValue("@procedureName", procedureName);

            var value = await command.ExecuteScalarAsync(cancellationToken);
            return value is null or DBNull ? null : value.ToString();
        }
        catch
        {
            return null;
        }
    }
}

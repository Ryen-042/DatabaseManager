using Microsoft.Data.SqlClient;

namespace DatabaseManager.Core.Services;

/// <summary>
/// Small connection-string helpers so callers (the WPF layer's ConnectionPickerWindow, for the
/// Databases tab) never need to reference Microsoft.Data.SqlClient directly - the WPF layer
/// talks to ADO.NET only through Core service interfaces, this included.
/// </summary>
public static class ConnectionStringHelper
{
    /// <summary>Returns a copy of the connection string with Initial Catalog set to <paramref name="databaseName"/>.</summary>
    public static string WithDatabase(string connectionString, string databaseName)
    {
        var builder = new SqlConnectionStringBuilder(connectionString)
        {
            InitialCatalog = databaseName
        };

        return builder.ConnectionString;
    }

    /// <summary>
    /// Returns a copy of the connection string with Initial Catalog set to
    /// <paramref name="fallbackDatabaseName"/> only if it isn't already set - used to give a
    /// connection string a valid database to connect to before listing sys.databases, without
    /// clobbering a database the caller already specified.
    /// </summary>
    public static string WithFallbackDatabase(string connectionString, string fallbackDatabaseName)
    {
        var builder = new SqlConnectionStringBuilder(connectionString);
        if (string.IsNullOrWhiteSpace(builder.InitialCatalog))
        {
            builder.InitialCatalog = fallbackDatabaseName;
        }

        return builder.ConnectionString;
    }
}

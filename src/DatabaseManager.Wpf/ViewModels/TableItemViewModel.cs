using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.ViewModels;

public sealed class TableItemViewModel(TableSchemaInfo table)
{
    public TableSchemaInfo Table { get; } = table;
    public string FullName => Table.FullName;
}

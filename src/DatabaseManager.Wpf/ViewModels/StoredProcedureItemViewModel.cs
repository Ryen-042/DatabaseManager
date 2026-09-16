using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.ViewModels;

public sealed class StoredProcedureItemViewModel(StoredProcedureSchemaInfo procedure)
{
    public StoredProcedureSchemaInfo Procedure { get; } = procedure;
    public string FullName => Procedure.FullName;
}

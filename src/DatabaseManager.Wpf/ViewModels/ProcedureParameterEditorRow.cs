namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// One editable row in the Procedure Runner's parameter grid. Plain mutable properties are
/// enough here - the DataGrid's own bindings write Value/SendAsNull directly on cell edit, and
/// nothing else needs to react to those changes while the grid is visible, so this doesn't need
/// to be an ObservableObject.
/// </summary>
public sealed class ProcedureParameterEditorRow
{
    public required string ParameterName { get; init; }

    public required string DataType { get; init; }

    public string? Value { get; set; }

    public bool SendAsNull { get; set; }

    public bool IsOutput { get; init; }
}

using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Core.Services.Schema;

/// <summary>
/// Pure editor-row-to-execution-parameter mapping and NULL handling for the Procedure Runner.
/// GetInputValue is the one piece of real logic here (deciding when a typed-in value becomes
/// DBNull.Value) and previously lived as a private method on StoredProcedureExecutionService,
/// which can't be unit tested without a live connection since every one of its public methods
/// opens one.
/// </summary>
public static class ProcedureParameterMapper
{
    public static StoredProcedureExecutionParameter ToExecutionParameter(string parameterName, string? value, bool sendAsNull, bool isOutput) => new()
    {
        Name = parameterName,
        Value = value,
        IsOutput = isOutput,
        IsInputOutput = false,
        SendAsNull = sendAsNull
    };

    /// <summary>
    /// The value actually sent to SQL Server for an input parameter: DBNull.Value when the
    /// user checked "send as NULL", or when the typed value is empty/whitespace (an untouched
    /// parameter shouldn't be sent as the literal string ""); otherwise the raw string value.
    /// </summary>
    public static object GetInputValue(StoredProcedureExecutionParameter parameter)
    {
        if (parameter.SendAsNull)
        {
            return DBNull.Value;
        }

        if (string.IsNullOrWhiteSpace(parameter.Value))
        {
            return DBNull.Value;
        }

        return parameter.Value;
    }
}

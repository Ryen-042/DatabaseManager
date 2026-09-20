using DatabaseManager.Core.Models.Schema;
using DatabaseManager.Core.Services.Schema;

namespace DatabaseManager.Tests.Services;

public sealed class ProcedureParameterMapperTests
{
    [Fact]
    public void ToExecutionParameter_CopiesAllFields()
    {
        var parameter = ProcedureParameterMapper.ToExecutionParameter("@id", "42", sendAsNull: false, isOutput: true);

        Assert.Equal("@id", parameter.Name);
        Assert.Equal("42", parameter.Value);
        Assert.True(parameter.IsOutput);
        Assert.False(parameter.IsInputOutput);
        Assert.False(parameter.SendAsNull);
    }

    [Fact]
    public void GetInputValue_SendAsNull_ReturnsDbNullRegardlessOfValue()
    {
        var parameter = new StoredProcedureExecutionParameter { Name = "@id", Value = "42", SendAsNull = true };

        Assert.Equal(DBNull.Value, ProcedureParameterMapper.GetInputValue(parameter));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void GetInputValue_EmptyOrWhitespaceValue_ReturnsDbNull(string? value)
    {
        var parameter = new StoredProcedureExecutionParameter { Name = "@id", Value = value, SendAsNull = false };

        Assert.Equal(DBNull.Value, ProcedureParameterMapper.GetInputValue(parameter));
    }

    [Fact]
    public void GetInputValue_NonEmptyValue_ReturnsTheRawStringValue()
    {
        var parameter = new StoredProcedureExecutionParameter { Name = "@id", Value = "42", SendAsNull = false };

        Assert.Equal("42", ProcedureParameterMapper.GetInputValue(parameter));
    }
}

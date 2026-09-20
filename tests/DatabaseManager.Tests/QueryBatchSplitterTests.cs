using DatabaseManager.Core.Services;

namespace DatabaseManager.Tests;

public sealed class QueryBatchSplitterTests
{
    [Fact]
    public void Split_NoGoSeparators_ReturnsSingleTrimmedBatch()
    {
        var sql = "  SELECT * FROM dbo.Users;  ";

        var batches = QueryBatchSplitter.Split(sql);

        Assert.Single(batches);
        Assert.Equal("SELECT * FROM dbo.Users;", batches[0]);
    }

    [Fact]
    public void Split_SsmsStyleScript_SplitsOnEachGoLine()
    {
        var sql =
            "USE [MVMASTER]\n" +
            "GO\n" +
            "SET ANSI_NULLS ON\n" +
            "GO\n" +
            "SET QUOTED_IDENTIFIER ON\n" +
            "GO\n" +
            "\n" +
            "CREATE PROC [dbo].[MyProc]\n" +
            "@fromDate datetime\n" +
            "AS\n" +
            "BEGIN\n" +
            "  SELECT 1;\n" +
            "END\n";

        var batches = QueryBatchSplitter.Split(sql);

        Assert.Equal(4, batches.Count);
        Assert.Equal("USE [MVMASTER]", batches[0]);
        Assert.Equal("SET ANSI_NULLS ON", batches[1]);
        Assert.Equal("SET QUOTED_IDENTIFIER ON", batches[2]);
        Assert.StartsWith("CREATE PROC [dbo].[MyProc]", batches[3]);
    }

    [Theory]
    [InlineData("GO")]
    [InlineData("go")]
    [InlineData("  GO  ")]
    [InlineData("GO 5")]
    public void Split_RecognizesGoVariants(string goLine)
    {
        var sql = $"SELECT 1;\n{goLine}\nSELECT 2;";

        var batches = QueryBatchSplitter.Split(sql);

        Assert.Equal(2, batches.Count);
        Assert.Equal("SELECT 1;", batches[0]);
        Assert.Equal("SELECT 2;", batches[1]);
    }

    [Fact]
    public void Split_BlankOrWhitespaceOnly_ReturnsEmpty()
    {
        Assert.Empty(QueryBatchSplitter.Split(""));
        Assert.Empty(QueryBatchSplitter.Split("   \n  "));
    }

    [Fact]
    public void Split_DoesNotTreatGoInsideAWordAsSeparator()
    {
        var sql = "SELECT * FROM Goods;";

        var batches = QueryBatchSplitter.Split(sql);

        Assert.Single(batches);
        Assert.Equal("SELECT * FROM Goods;", batches[0]);
    }

    [Fact]
    public void SplitThenExtractParameterNames_SsmsStyleCreateProcScript_ReturnsNoParameters()
    {
        // Reproduces the real-world "Script Stored Procedure as CREATE" shape: USE/SET/GO
        // boilerplate ahead of the actual CREATE PROC batch. Splitting first means each
        // ExtractParameterNames call sees the CREATE PROC batch in isolation, so its
        // @fromDate/@toDate parameter declarations are correctly recognized as part of the
        // routine definition rather than query parameters to prompt for.
        var sql =
            "USE [MVMASTER]\n" +
            "GO\n" +
            "SET ANSI_NULLS ON\n" +
            "GO\n" +
            "\n" +
            "Create PROC [dbo].[GetFrxEODFileTest1]\n" +
            "@fromDate Datetime , @toDate datetime\n" +
            "As\n" +
            "Begin\n" +
            "  select 1;\n" +
            "End\n";

        var parameterNames = QueryBatchSplitter.Split(sql)
            .SelectMany(QueryOutputModeParser.ExtractParameterNames)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        Assert.Empty(parameterNames);
    }
}

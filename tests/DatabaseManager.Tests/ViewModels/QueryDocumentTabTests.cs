using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class QueryDocumentTabTests
{
    [Fact]
    public void BeginRename_StagesCurrentTitleAndEntersRenamingMode()
    {
        var document = new QueryDocumentTab("Query 1");

        document.BeginRenameCommand.Execute(null);

        Assert.True(document.IsRenaming);
        Assert.Equal("Query 1", document.PendingTitle);
    }

    [Fact]
    public void CommitRename_TrimmedNonEmptyPendingTitle_AppliesItAndExitsRenamingMode()
    {
        var document = new QueryDocumentTab("Query 1");
        document.BeginRenameCommand.Execute(null);
        document.PendingTitle = "  Sales Report  ";

        document.CommitRenameCommand.Execute(null);

        Assert.False(document.IsRenaming);
        Assert.Equal("Sales Report", document.Title);
    }

    [Fact]
    public void CommitRename_EmptyOrWhitespacePendingTitle_KeepsOldTitle()
    {
        var document = new QueryDocumentTab("Query 1");
        document.BeginRenameCommand.Execute(null);
        document.PendingTitle = "   ";

        document.CommitRenameCommand.Execute(null);

        Assert.False(document.IsRenaming);
        Assert.Equal("Query 1", document.Title);
    }

    [Fact]
    public void CommitRename_WhenNotRenaming_IsANoOp()
    {
        var document = new QueryDocumentTab("Query 1");
        document.PendingTitle = "Ignored";

        document.CommitRenameCommand.Execute(null);

        Assert.Equal("Query 1", document.Title);
    }

    [Fact]
    public void CancelRename_DiscardsPendingTitleAndExitsRenamingMode()
    {
        var document = new QueryDocumentTab("Query 1");
        document.BeginRenameCommand.Execute(null);
        document.PendingTitle = "Discarded";

        document.CancelRenameCommand.Execute(null);

        Assert.False(document.IsRenaming);
        Assert.Equal("Query 1", document.Title);
    }

    [Fact]
    public void SetColor_NonEmptyHex_SetsColor()
    {
        var document = new QueryDocumentTab("Query 1");

        document.SetColorCommand.Execute("#E5484D");

        Assert.Equal("#E5484D", document.Color);
    }

    [Fact]
    public void SetColor_NullOrWhitespace_ClearsColor()
    {
        var document = new QueryDocumentTab("Query 1");
        document.SetColorCommand.Execute("#E5484D");

        document.SetColorCommand.Execute(null);

        Assert.Null(document.Color);
    }
}

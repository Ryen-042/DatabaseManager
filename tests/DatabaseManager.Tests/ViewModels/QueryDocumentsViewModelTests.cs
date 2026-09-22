using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class QueryDocumentsViewModelTests
{
    private sealed class TestHarness
    {
        public string EditorText { get; set; } = string.Empty;
        public List<QueryDocumentTab> Activations { get; } = new();
        public bool ConfirmCloseResult { get; set; } = true;
        public List<QueryDocumentTab> CloseConfirmationPrompts { get; } = new();

        public QueryDocumentsViewModel CreateViewModel() => new(
            getEditorText: () => EditorText,
            activateDocument: Activations.Add,
            confirmCloseDirtyDocument: document =>
            {
                CloseConfirmationPrompts.Add(document);
                return ConfirmCloseResult;
            });
    }

    [Fact]
    public void Constructor_StartsWithOneDocumentSelected()
    {
        var vm = new TestHarness().CreateViewModel();

        Assert.Single(vm.Documents);
        Assert.Same(vm.Documents[0], vm.SelectedDocument);
        Assert.Equal("Query 1", vm.SelectedDocument!.Title);
    }

    [Fact]
    public void NewDocument_AddsAndSelectsANewNumberedDocument()
    {
        var vm = new TestHarness().CreateViewModel();

        vm.NewDocumentCommand.Execute(null);

        Assert.Equal(2, vm.Documents.Count);
        Assert.Equal("Query 2", vm.SelectedDocument!.Title);
        Assert.Same(vm.Documents[1], vm.SelectedDocument);
    }

    [Fact]
    public void SelectingADocument_SavesOutgoingTextAndActivatesIncoming()
    {
        var harness = new TestHarness { EditorText = "SELECT 1;" };
        var vm = harness.CreateViewModel();
        var first = vm.Documents[0];

        vm.NewDocumentCommand.Execute(null);
        var second = vm.SelectedDocument!;

        Assert.Equal("SELECT 1;", first.SqlText);
        Assert.Contains(second, harness.Activations);

        harness.EditorText = "SELECT 2;";
        vm.SelectedDocument = first;

        Assert.Equal("SELECT 2;", second.SqlText);
        Assert.Same(first, harness.Activations[^1]);
    }

    [Fact]
    public void SelectingTheAlreadySelectedDocument_DoesNothing()
    {
        var harness = new TestHarness();
        var vm = harness.CreateViewModel();
        harness.Activations.Clear();

        vm.SelectedDocument = vm.SelectedDocument;

        Assert.Empty(harness.Activations);
    }

    [Fact]
    public void CloseDocument_CleanDocument_ClosesWithoutPrompting()
    {
        var harness = new TestHarness();
        var vm = harness.CreateViewModel();
        vm.NewDocumentCommand.Execute(null);
        var toClose = vm.SelectedDocument!;

        vm.CloseDocumentCommand.Execute(toClose);

        Assert.DoesNotContain(toClose, vm.Documents);
        Assert.Empty(harness.CloseConfirmationPrompts);
    }

    [Fact]
    public void CloseDocument_DirtyDocument_PromptsForConfirmation()
    {
        var harness = new TestHarness();
        var vm = harness.CreateViewModel();
        vm.NewDocumentCommand.Execute(null);
        var toClose = vm.SelectedDocument!;
        toClose.IsDirty = true;

        vm.CloseDocumentCommand.Execute(toClose);

        Assert.Contains(toClose, harness.CloseConfirmationPrompts);
        Assert.DoesNotContain(toClose, vm.Documents);
    }

    [Fact]
    public void CloseDocument_DirtyDocumentConfirmationDeclined_KeepsItOpen()
    {
        var harness = new TestHarness { ConfirmCloseResult = false };
        var vm = harness.CreateViewModel();
        vm.NewDocumentCommand.Execute(null);
        var toClose = vm.SelectedDocument!;
        toClose.IsDirty = true;

        vm.CloseDocumentCommand.Execute(toClose);

        Assert.Contains(toClose, vm.Documents);
    }

    [Fact]
    public void CloseDocument_LastRemainingDocument_ReplacesItWithAFreshOne()
    {
        var vm = new TestHarness().CreateViewModel();
        var only = vm.Documents[0];

        vm.CloseDocumentCommand.Execute(only);

        Assert.Single(vm.Documents);
        Assert.NotSame(only, vm.Documents[0]);
        Assert.Same(vm.Documents[0], vm.SelectedDocument);
    }

    [Fact]
    public void CloseDocument_NotTheSelectedOne_LeavesSelectionUnchanged()
    {
        var vm = new TestHarness().CreateViewModel();
        var first = vm.Documents[0];
        vm.NewDocumentCommand.Execute(null);
        var second = vm.SelectedDocument!;
        vm.NewDocumentCommand.Execute(null);
        var third = vm.SelectedDocument!;

        vm.CloseDocumentCommand.Execute(second);

        Assert.DoesNotContain(second, vm.Documents);
        Assert.Same(third, vm.SelectedDocument);
        Assert.Contains(first, vm.Documents);
    }

    [Fact]
    public void CloseDocument_NullParameter_ClosesCurrentlySelectedDocument()
    {
        var vm = new TestHarness().CreateViewModel();
        vm.NewDocumentCommand.Execute(null);
        var selected = vm.SelectedDocument!;

        vm.CloseDocumentCommand.Execute(null);

        Assert.DoesNotContain(selected, vm.Documents);
    }

    [Fact]
    public void Pinning_MovesDocumentToEndOfThePinnedGroup()
    {
        var vm = new TestHarness().CreateViewModel();
        var first = vm.Documents[0];
        vm.NewDocumentCommand.Execute(null);
        var second = vm.SelectedDocument!;
        vm.NewDocumentCommand.Execute(null);
        var third = vm.SelectedDocument!;

        first.IsPinned = true;
        third.IsPinned = true;

        Assert.Equal(new[] { first, third, second }, vm.Documents);
    }

    [Fact]
    public void Unpinning_MovesDocumentToStartOfTheUnpinnedGroup()
    {
        var vm = new TestHarness().CreateViewModel();
        var first = vm.Documents[0];
        vm.NewDocumentCommand.Execute(null);
        var second = vm.SelectedDocument!;
        vm.NewDocumentCommand.Execute(null);
        var third = vm.SelectedDocument!;

        first.IsPinned = true;
        second.IsPinned = true;
        // Order is now: first, second, third (both pinned docs at the front).
        first.IsPinned = false;

        Assert.Equal(new[] { second, first, third }, vm.Documents);
    }

    [Fact]
    public void ClosingAPinnedDocument_StopsItFromAffectingFurtherReordering()
    {
        var vm = new TestHarness().CreateViewModel();
        var first = vm.Documents[0];
        vm.NewDocumentCommand.Execute(null);
        var second = vm.SelectedDocument!;
        first.IsPinned = true;

        vm.CloseDocumentCommand.Execute(first);
        // If the closed document's PropertyChanged subscription weren't removed, flipping
        // IsPinned on it afterward would throw (or silently reorder a collection it's no longer
        // in) instead of being inert.
        first.IsPinned = false;

        Assert.DoesNotContain(first, vm.Documents);
        Assert.Single(vm.Documents);
        Assert.Same(second, vm.Documents[0]);
    }
}

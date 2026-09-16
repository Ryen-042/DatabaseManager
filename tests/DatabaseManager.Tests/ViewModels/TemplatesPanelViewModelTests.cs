using DatabaseManager.Core.Services;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Tests.ViewModels;

public sealed class TemplatesPanelViewModelTests
{
    private static (TemplatesPanelViewModel ViewModel, string TempFile, List<string> Statuses) Create(
        string sqlToSave = "SELECT 1;",
        bool confirmDelete = true,
        List<(string Name, string Sql)>? activations = null)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"dbm-templates-{Guid.NewGuid():N}.json");
        var store = new TemplateStoreService(tempFile);
        var statuses = new List<string>();

        var vm = new TemplatesPanelViewModel(
            store,
            getCurrentSqlText: () => sqlToSave,
            onTemplateActivated: (name, sql) => activations?.Add((name, sql)),
            confirmDelete: _ => confirmDelete,
            setStatus: statuses.Add);

        return (vm, tempFile, statuses);
    }

    [Fact]
    public async Task RefreshAsync_PopulatesTemplatesFromStore()
    {
        var (vm, tempFile, _) = Create();
        try
        {
            var store = new TemplateStoreService(tempFile);
            await store.SaveAsync(new() { Name = "Existing", Sql = "SELECT 2;" }, CancellationToken.None);

            await vm.RefreshAsync();

            Assert.Single(vm.Templates);
            Assert.Equal("Existing", vm.Templates[0].Name);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task SaveCommand_ValidNameAndSql_SavesAndAddsToList()
    {
        var (vm, tempFile, statuses) = Create(sqlToSave: "SELECT * FROM Foo;");
        try
        {
            vm.NewTemplateName = "MyTemplate";

            await vm.SaveCommand.ExecuteAsync(null);

            Assert.Single(vm.Templates);
            Assert.Equal("MyTemplate", vm.Templates[0].Name);
            Assert.Equal("SELECT * FROM Foo;", vm.Templates[0].Sql);
            Assert.Contains("Template 'MyTemplate' saved.", statuses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task SaveCommand_EmptyName_DoesNotSave()
    {
        var (vm, tempFile, statuses) = Create();
        try
        {
            vm.NewTemplateName = "   ";

            await vm.SaveCommand.ExecuteAsync(null);

            Assert.Empty(vm.Templates);
            Assert.Contains("Template name is required.", statuses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task SaveCommand_EmptySql_DoesNotSave()
    {
        var (vm, tempFile, statuses) = Create(sqlToSave: "   ");
        try
        {
            vm.NewTemplateName = "MyTemplate";

            await vm.SaveCommand.ExecuteAsync(null);

            Assert.Empty(vm.Templates);
            Assert.Contains("Cannot save an empty SQL template.", statuses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task DeleteCommand_NoSelection_DoesNothing()
    {
        var (vm, tempFile, statuses) = Create();
        try
        {
            await vm.DeleteCommand.ExecuteAsync(null);

            Assert.Contains("Select a template to delete.", statuses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task DeleteCommand_ConfirmationDeclined_DoesNotDelete()
    {
        var (vm, tempFile, _) = Create(confirmDelete: false);
        try
        {
            vm.NewTemplateName = "ToDelete";
            await vm.SaveCommand.ExecuteAsync(null);
            vm.SelectedTemplate = vm.Templates[0];

            await vm.DeleteCommand.ExecuteAsync(null);

            Assert.Single(vm.Templates);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task DeleteCommand_Confirmed_RemovesTemplate()
    {
        var (vm, tempFile, statuses) = Create();
        try
        {
            vm.NewTemplateName = "ToDelete";
            await vm.SaveCommand.ExecuteAsync(null);
            vm.SelectedTemplate = vm.Templates[0];

            await vm.DeleteCommand.ExecuteAsync(null);

            Assert.Empty(vm.Templates);
            Assert.Contains("Template 'ToDelete' deleted.", statuses);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task ActivateSelected_NoSelection_DoesNotInvokeCallback()
    {
        var activations = new List<(string Name, string Sql)>();
        var (vm, tempFile, _) = Create(activations: activations);
        try
        {
            vm.ActivateSelected();

            Assert.Empty(activations);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Fact]
    public async Task ActivateSelected_WithSelection_InvokesCallbackAndPrefillsName()
    {
        var activations = new List<(string Name, string Sql)>();
        var (vm, tempFile, _) = Create(sqlToSave: "SELECT 42;", activations: activations);
        try
        {
            vm.NewTemplateName = "Activated";
            await vm.SaveCommand.ExecuteAsync(null);
            vm.SelectedTemplate = vm.Templates[0];
            vm.NewTemplateName = string.Empty; // simulate the user having cleared/typed over it

            vm.ActivateSelected();

            Assert.Single(activations);
            Assert.Equal(("Activated", "SELECT 42;"), activations[0]);
            Assert.Equal("Activated", vm.NewTemplateName);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}

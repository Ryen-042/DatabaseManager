using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Tests.Commands;

public sealed class AppCommandRegistryTests
{
    private static AppCommandDescriptor MakeDescriptor(string id, string category = "Category", string name = "Name") =>
        new()
        {
            Id = id,
            DisplayName = name,
            Category = category,
            Command = new RelayCommand(() => { })
        };

    [Fact]
    public void Register_AddsCommand_VisibleInCommands()
    {
        var registry = new AppCommandRegistry();

        registry.Register(MakeDescriptor("cmd.one"));

        Assert.Single(registry.Commands);
        Assert.Equal("cmd.one", registry.Commands[0].Id);
    }

    [Fact]
    public void Register_DuplicateId_Throws()
    {
        var registry = new AppCommandRegistry();
        registry.Register(MakeDescriptor("cmd.one"));

        Assert.Throws<InvalidOperationException>(() => registry.Register(MakeDescriptor("cmd.one")));
    }

    [Fact]
    public void Unregister_RemovesCommand()
    {
        var registry = new AppCommandRegistry();
        registry.Register(MakeDescriptor("cmd.one"));

        registry.Unregister("cmd.one");

        Assert.Empty(registry.Commands);
    }

    [Fact]
    public void Unregister_UnknownId_DoesNotThrowOrRaiseEvent()
    {
        var registry = new AppCommandRegistry();
        var raised = false;
        registry.CommandsChanged += (_, _) => raised = true;

        registry.Unregister("does.not.exist");

        Assert.False(raised);
    }

    [Fact]
    public void Register_RaisesCommandsChanged()
    {
        var registry = new AppCommandRegistry();
        var raiseCount = 0;
        registry.CommandsChanged += (_, _) => raiseCount++;

        registry.Register(MakeDescriptor("cmd.one"));

        Assert.Equal(1, raiseCount);
    }

    [Fact]
    public void Commands_AreOrderedByCategoryThenDisplayName()
    {
        var registry = new AppCommandRegistry();
        registry.Register(MakeDescriptor("cmd.b", category: "Beta", name: "Zeta"));
        registry.Register(MakeDescriptor("cmd.a1", category: "Alpha", name: "Zeta"));
        registry.Register(MakeDescriptor("cmd.a2", category: "Alpha", name: "Alpha"));

        var ordered = registry.Commands.Select(c => c.Id).ToList();

        Assert.Equal(new[] { "cmd.a2", "cmd.a1", "cmd.b" }, ordered);
    }
}

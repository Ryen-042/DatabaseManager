namespace DatabaseManager.Wpf.Commands;

public interface ICommandRegistry
{
    void Register(AppCommandDescriptor descriptor);

    void Unregister(string id);

    IReadOnlyList<AppCommandDescriptor> Commands { get; }

    event EventHandler? CommandsChanged;
}

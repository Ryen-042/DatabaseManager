namespace DatabaseManager.Wpf.Commands;

public sealed class AppCommandRegistry : ICommandRegistry
{
    private readonly Dictionary<string, AppCommandDescriptor> _commandsById = new(StringComparer.Ordinal);

    public event EventHandler? CommandsChanged;

    public IReadOnlyList<AppCommandDescriptor> Commands => _commandsById.Values
        .OrderBy(c => c.Category, StringComparer.OrdinalIgnoreCase)
        .ThenBy(c => c.DisplayName, StringComparer.OrdinalIgnoreCase)
        .ToList();

    public void Register(AppCommandDescriptor descriptor)
    {
        if (_commandsById.ContainsKey(descriptor.Id))
        {
            throw new InvalidOperationException($"A command with id '{descriptor.Id}' is already registered.");
        }

        _commandsById[descriptor.Id] = descriptor;
        CommandsChanged?.Invoke(this, EventArgs.Empty);
    }

    public void Unregister(string id)
    {
        if (_commandsById.Remove(id))
        {
            CommandsChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}

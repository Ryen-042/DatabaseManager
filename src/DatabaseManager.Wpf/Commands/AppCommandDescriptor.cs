using System.Windows.Input;

namespace DatabaseManager.Wpf.Commands;

/// <summary>
/// A single app-level action, registered once and shared by the top menu, the command
/// palette, the shortcuts help panel, and the window's keyboard InputBindings - so a
/// shortcut can never drift out of sync between where it's shown and what it does.
/// </summary>
public sealed class AppCommandDescriptor
{
    public required string Id { get; init; }

    public required string DisplayName { get; init; }

    public required string Category { get; init; }

    public string? IconKey { get; init; }

    public KeyGesture? Gesture { get; init; }

    public required ICommand Command { get; init; }

    public IReadOnlyList<string> Keywords { get; init; } = Array.Empty<string>();
}

using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace DatabaseManager.Wpf.ViewModels;

public enum ToastKind
{
    Info,
    Success,
    Warning,
    Error
}

public sealed partial class ToastNotificationViewModel : ObservableObject
{
    public required string Message { get; init; }

    public required ToastKind Kind { get; init; }

    public event EventHandler? DismissRequested;

    [RelayCommand]
    private void Dismiss() => DismissRequested?.Invoke(this, EventArgs.Empty);
}

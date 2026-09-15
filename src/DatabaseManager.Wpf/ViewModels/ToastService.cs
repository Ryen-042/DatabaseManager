using System.Collections.ObjectModel;
using System.Windows.Threading;

namespace DatabaseManager.Wpf.ViewModels;

public sealed class ToastService : IToastService
{
    private static readonly TimeSpan DefaultDuration = TimeSpan.FromSeconds(5);

    private readonly ObservableCollection<ToastNotificationViewModel> _toasts = new();
    private readonly Dispatcher _dispatcher;

    public ToastService(Dispatcher dispatcher)
    {
        _dispatcher = dispatcher;
        Toasts = new ReadOnlyObservableCollection<ToastNotificationViewModel>(_toasts);
    }

    public ReadOnlyObservableCollection<ToastNotificationViewModel> Toasts { get; }

    public void Show(string message, ToastKind kind = ToastKind.Info, TimeSpan? duration = null)
    {
        var toast = new ToastNotificationViewModel { Message = message, Kind = kind };
        toast.DismissRequested += Toast_DismissRequested;
        _toasts.Add(toast);

        var timer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
        {
            Interval = duration ?? DefaultDuration
        };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            Remove(toast);
        };
        timer.Start();
    }

    private void Toast_DismissRequested(object? sender, EventArgs e)
    {
        if (sender is ToastNotificationViewModel toast)
        {
            Remove(toast);
        }
    }

    private void Remove(ToastNotificationViewModel toast)
    {
        toast.DismissRequested -= Toast_DismissRequested;
        _toasts.Remove(toast);
    }
}

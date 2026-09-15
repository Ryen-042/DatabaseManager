using System.Collections.ObjectModel;

namespace DatabaseManager.Wpf.ViewModels;

public interface IToastService
{
    ReadOnlyObservableCollection<ToastNotificationViewModel> Toasts { get; }

    void Show(string message, ToastKind kind = ToastKind.Info, TimeSpan? duration = null);
}

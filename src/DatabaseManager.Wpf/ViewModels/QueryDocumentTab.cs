using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// One open query buffer in the Query tab's document strip: its own SQL text, dirty flag, and
/// last execution results. Execution itself is NOT per-document - QueryDocumentViewModel is a
/// single shared engine (only one query runs at a time, by design; see CLAUDE.md/plan.md for
/// why), so this class is pure state, not a second execution pipeline. Title/IsDirty are
/// observable because the document strip binds to them directly; SqlText and the stored results
/// are plain properties copied to/from the shared editor and results panel on tab switch.
///
/// IsPinned/Color/rename are also plain per-document state with no cross-document side effects
/// of their own - QueryDocumentsViewModel listens for IsPinned's PropertyChanged to reorder the
/// document strip (pinned documents sort to the front), since ordering is a collection-level
/// concern this class doesn't own itself.
/// </summary>
public sealed partial class QueryDocumentTab : ObservableObject
{
    public QueryDocumentTab(string title)
    {
        Title = title;
    }

    [ObservableProperty]
    private string _title;

    [ObservableProperty]
    private bool _isDirty;

    [ObservableProperty]
    private bool _isPinned;

    /// <summary>Hex color string (e.g. "#E5484D") for the tab's color indicator strip, or null for none.</summary>
    [ObservableProperty]
    private string? _color;

    [ObservableProperty]
    private bool _isRenaming;

    /// <summary>Staged edit for the title TextBox while IsRenaming is true - only copied to Title on commit.</summary>
    [ObservableProperty]
    private string _pendingTitle = string.Empty;

    public string SqlText { get; set; } = string.Empty;

    public IReadOnlyList<QueryResultSet>? LastResultSets { get; set; }

    public string? LastResultsSummary { get; set; }

    [RelayCommand]
    private void BeginRename()
    {
        PendingTitle = Title;
        IsRenaming = true;
    }

    [RelayCommand]
    private void CommitRename()
    {
        if (!IsRenaming)
        {
            return;
        }

        var trimmed = PendingTitle.Trim();
        if (trimmed.Length > 0)
        {
            Title = trimmed;
        }

        IsRenaming = false;
    }

    [RelayCommand]
    private void CancelRename() => IsRenaming = false;

    [RelayCommand]
    private void SetColor(string? colorHex) => Color = string.IsNullOrWhiteSpace(colorHex) ? null : colorHex;
}

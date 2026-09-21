using CommunityToolkit.Mvvm.ComponentModel;
using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// One open query buffer in the Query tab's document strip: its own SQL text, dirty flag, and
/// last execution results. Execution itself is NOT per-document - QueryDocumentViewModel is a
/// single shared engine (only one query runs at a time, by design; see CLAUDE.md/plan.md for
/// why), so this class is pure state, not a second execution pipeline. Title/IsDirty are
/// observable because the document strip binds to them directly; SqlText and the stored results
/// are plain properties copied to/from the shared editor and results panel on tab switch.
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

    public string SqlText { get; set; } = string.Empty;

    public IReadOnlyList<QueryResultSet>? LastResultSets { get; set; }

    public string? LastResultsSummary { get; set; }
}

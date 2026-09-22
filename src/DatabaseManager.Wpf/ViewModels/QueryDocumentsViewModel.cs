using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Owns the Query tab's document strip: which QueryDocumentTabs are open and which is selected.
/// Execution (QueryDocumentViewModel) stays a single shared engine bound to whichever document
/// is currently selected - switching the selection here saves the outgoing document's text out
/// of the shared editor and loads the incoming one's back in, via callbacks MainWindow supplies,
/// the same pattern every other extracted ViewModel uses for things outside its own concern.
///
/// Also reorders Documents when a document's IsPinned flips (pinned documents sort to the front)
/// - QueryDocumentTab owns the flag itself but not the reordering, since that's a collection-level
/// concern; this class listens for the PropertyChanged notification instead.
/// </summary>
public sealed partial class QueryDocumentsViewModel : ObservableObject
{
    private readonly Func<string> _getEditorText;
    private readonly Action<QueryDocumentTab> _activateDocument;
    private readonly Func<QueryDocumentTab, bool> _confirmCloseDirtyDocument;
    private int _nextDocumentNumber = 1;
    private QueryDocumentTab? _selectedDocument;

    public QueryDocumentsViewModel(
        Func<string> getEditorText,
        Action<QueryDocumentTab> activateDocument,
        Func<QueryDocumentTab, bool> confirmCloseDirtyDocument)
    {
        _getEditorText = getEditorText;
        _activateDocument = activateDocument;
        _confirmCloseDirtyDocument = confirmCloseDirtyDocument;

        var first = CreateDocument();
        Documents.Add(first);
        _selectedDocument = first;
    }

    public ObservableCollection<QueryDocumentTab> Documents { get; } = new();

    /// <summary>
    /// Hand-written rather than [ObservableProperty] so the save-outgoing/activate-incoming
    /// sequencing around the shared editor and results panel is explicit and ordered, rather
    /// than depending on exactly which generated On*Changing/On*Changed overload the source
    /// generator emits.
    /// </summary>
    public QueryDocumentTab? SelectedDocument
    {
        get => _selectedDocument;
        set
        {
            if (ReferenceEquals(_selectedDocument, value))
            {
                return;
            }

            if (_selectedDocument is not null)
            {
                _selectedDocument.SqlText = _getEditorText();
            }

            SetProperty(ref _selectedDocument, value);

            if (value is not null)
            {
                _activateDocument(value);
            }
        }
    }

    private QueryDocumentTab CreateDocument()
    {
        var document = new QueryDocumentTab($"Query {_nextDocumentNumber++}");
        document.PropertyChanged += OnDocumentPropertyChanged;
        return document;
    }

    private void OnDocumentPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(QueryDocumentTab.IsPinned) && sender is QueryDocumentTab document)
        {
            ReorderAfterPinChange(document);
        }
    }

    /// <summary>
    /// Removing then reinserting at "count of remaining pinned documents" places a newly-pinned
    /// document at the end of the pinned group, and a newly-unpinned document at the start of the
    /// unpinned group - the same formula works for both directions.
    /// </summary>
    private void ReorderAfterPinChange(QueryDocumentTab document)
    {
        var currentIndex = Documents.IndexOf(document);
        if (currentIndex < 0)
        {
            return;
        }

        Documents.RemoveAt(currentIndex);
        var insertIndex = Documents.TakeWhile(d => d.IsPinned).Count();
        Documents.Insert(insertIndex, document);
    }

    [RelayCommand]
    private void NewDocument()
    {
        var document = CreateDocument();
        Documents.Add(document);
        SelectedDocument = document;
    }

    [RelayCommand]
    private void CloseDocument(QueryDocumentTab? document)
    {
        document ??= SelectedDocument;
        if (document is null)
        {
            return;
        }

        if (document.IsDirty && !_confirmCloseDirtyDocument(document))
        {
            return;
        }

        var index = Documents.IndexOf(document);
        if (index < 0)
        {
            return;
        }

        document.PropertyChanged -= OnDocumentPropertyChanged;

        var wasSelected = ReferenceEquals(SelectedDocument, document);

        if (wasSelected)
        {
            // Clear the selection first (without triggering a save-back into the document
            // we're about to remove) so the property setter above doesn't try to persist
            // editor text into a document that's going away.
            _selectedDocument = null;
        }

        Documents.RemoveAt(index);

        if (Documents.Count == 0)
        {
            var replacement = CreateDocument();
            Documents.Add(replacement);
            SelectedDocument = replacement;
            return;
        }

        if (wasSelected)
        {
            var newIndex = Math.Min(index, Documents.Count - 1);
            SelectedDocument = Documents[newIndex];
        }
    }
}

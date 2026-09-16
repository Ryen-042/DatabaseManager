using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DatabaseManager.Core.Models;
using DatabaseManager.Core.Services;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>
/// Backs the Query Templates panel (Schema Assistant sidebar). MainWindow still owns the SQL
/// editor and tab navigation, so this reaches back into it via constructor-supplied callbacks
/// for the operations that cross that boundary - reading the SQL to save, loading a template
/// into the editor, confirming a delete, and reporting status - rather than this ViewModel
/// reaching into named XAML elements it doesn't own.
/// </summary>
public sealed partial class TemplatesPanelViewModel : ObservableObject
{
    private readonly ITemplateStoreService _templateStoreService;
    private readonly Func<string> _getCurrentSqlText;
    private readonly Action<string, string> _onTemplateActivated;
    private readonly Func<string, bool> _confirmDelete;
    private readonly Action<string> _setStatus;

    public TemplatesPanelViewModel(
        ITemplateStoreService templateStoreService,
        Func<string> getCurrentSqlText,
        Action<string, string> onTemplateActivated,
        Func<string, bool> confirmDelete,
        Action<string> setStatus)
    {
        _templateStoreService = templateStoreService;
        _getCurrentSqlText = getCurrentSqlText;
        _onTemplateActivated = onTemplateActivated;
        _confirmDelete = confirmDelete;
        _setStatus = setStatus;
    }

    public ObservableCollection<TemplateItemViewModel> Templates { get; } = new();

    [ObservableProperty]
    private string _newTemplateName = string.Empty;

    [ObservableProperty]
    private TemplateItemViewModel? _selectedTemplate;

    public async Task RefreshAsync()
    {
        var templates = await _templateStoreService.GetAllAsync(CancellationToken.None);

        Templates.Clear();
        foreach (var template in templates)
        {
            Templates.Add(new TemplateItemViewModel(template));
        }
    }

    [RelayCommand]
    private async Task Refresh()
    {
        await RefreshAsync();
        _setStatus("Templates refreshed.");
    }

    [RelayCommand]
    private async Task Save()
    {
        var name = NewTemplateName.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            _setStatus("Template name is required.");
            return;
        }

        var sql = _getCurrentSqlText();
        if (string.IsNullOrWhiteSpace(sql))
        {
            _setStatus("Cannot save an empty SQL template.");
            return;
        }

        await _templateStoreService.SaveAsync(new QueryTemplate { Name = name, Sql = sql }, CancellationToken.None);
        await RefreshAsync();
        _setStatus($"Template '{name}' saved.");
    }

    [RelayCommand]
    private async Task Delete()
    {
        if (SelectedTemplate is null)
        {
            _setStatus("Select a template to delete.");
            return;
        }

        var name = SelectedTemplate.Name;
        if (!_confirmDelete(name))
        {
            return;
        }

        await _templateStoreService.DeleteAsync(name, CancellationToken.None);
        await RefreshAsync();
        _setStatus($"Template '{name}' deleted.");
    }

    public void ActivateSelected()
    {
        if (SelectedTemplate is null)
        {
            return;
        }

        // Matches the original double-click behavior: pre-fill the name field with the
        // loaded template's name so re-saving under the same name is a single click.
        NewTemplateName = SelectedTemplate.Name;
        _onTemplateActivated(SelectedTemplate.Name, SelectedTemplate.Sql);
        _setStatus($"Template '{SelectedTemplate.Name}' loaded into editor.");
    }
}

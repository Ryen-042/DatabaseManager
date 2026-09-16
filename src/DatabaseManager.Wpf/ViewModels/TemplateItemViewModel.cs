using DatabaseManager.Core.Models;

namespace DatabaseManager.Wpf.ViewModels;

/// <summary>Thin bindable wrapper around a <see cref="QueryTemplate"/> for display in the templates list.</summary>
public sealed class TemplateItemViewModel(QueryTemplate template)
{
    public QueryTemplate Template { get; } = template;

    public string Name => Template.Name;

    public string Sql => Template.Sql;
}

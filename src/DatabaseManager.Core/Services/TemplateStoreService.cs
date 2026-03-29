using System.Text.Json;
using DatabaseManager.Core.Models;

namespace DatabaseManager.Core.Services;

public sealed class TemplateStoreService : ITemplateStoreService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private readonly string _filePath;

    public TemplateStoreService(string? filePath = null)
    {
        var appData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DatabaseManager");

        Directory.CreateDirectory(appData);
        _filePath = filePath ?? Path.Combine(appData, "query-templates.json");
    }

    public async Task<IReadOnlyList<QueryTemplate>> GetAllAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_filePath))
        {
            return Array.Empty<QueryTemplate>();
        }

        await using var stream = File.OpenRead(_filePath);
        var items = await JsonSerializer.DeserializeAsync<List<QueryTemplate>>(stream, JsonOptions, cancellationToken);

        if (items is null)
        {
            return Array.Empty<QueryTemplate>();
        }

        return items
            .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task SaveAsync(QueryTemplate template, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(template.Name))
        {
            throw new ArgumentException("Template name is required.", nameof(template));
        }

        var templates = (await GetAllAsync(cancellationToken)).ToList();
        var existingIndex = templates.FindIndex(x => x.Name.Equals(template.Name, StringComparison.OrdinalIgnoreCase));
        var now = DateTimeOffset.UtcNow;

        if (existingIndex >= 0)
        {
            var existing = templates[existingIndex];
            templates[existingIndex] = new QueryTemplate
            {
                Name = template.Name.Trim(),
                Sql = template.Sql,
                CreatedAtUtc = existing.CreatedAtUtc,
                UpdatedAtUtc = now
            };
        }
        else
        {
            templates.Add(new QueryTemplate
            {
                Name = template.Name.Trim(),
                Sql = template.Sql,
                CreatedAtUtc = now,
                UpdatedAtUtc = now
            });
        }

        await WriteAllAsync(templates, cancellationToken);
    }

    public async Task DeleteAsync(string templateName, CancellationToken cancellationToken)
    {
        var templates = (await GetAllAsync(cancellationToken)).ToList();
        templates.RemoveAll(x => x.Name.Equals(templateName, StringComparison.OrdinalIgnoreCase));
        await WriteAllAsync(templates, cancellationToken);
    }

    private async Task WriteAllAsync(IReadOnlyCollection<QueryTemplate> templates, CancellationToken cancellationToken)
    {
        await using var stream = File.Create(_filePath);
        await JsonSerializer.SerializeAsync(stream, templates, JsonOptions, cancellationToken);
    }
}

using DatabaseManager.Core.Models;

namespace DatabaseManager.Core.Services;

public interface ITemplateStoreService
{
    Task<IReadOnlyList<QueryTemplate>> GetAllAsync(CancellationToken cancellationToken);

    Task SaveAsync(QueryTemplate template, CancellationToken cancellationToken);

    Task DeleteAsync(string templateName, CancellationToken cancellationToken);
}

using DatabaseManager.Core.Models.Schema;

namespace DatabaseManager.Wpf.SqlSuggestions;

public interface ISqlSuggestionEngine
{
    IReadOnlyList<string> GetSuggestions(SqlSuggestionRequest request);
}

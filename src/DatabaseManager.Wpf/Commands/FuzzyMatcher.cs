namespace DatabaseManager.Wpf.Commands;

/// <summary>
/// Minimal subsequence fuzzy matcher for the command palette: every character of the query
/// must appear in the target in order (not necessarily contiguous), scored so that longer
/// consecutive runs and earlier matches rank higher - the same rough heuristic VS Code's
/// command palette / "Quick Open" use.
/// </summary>
public static class FuzzyMatcher
{
    /// <summary>
    /// Returns a match score (higher is better), or null if <paramref name="query"/> is not
    /// a subsequence of <paramref name="target"/>. An empty query matches everything with a
    /// score of 0.
    /// </summary>
    public static int? Score(string query, string target)
    {
        if (string.IsNullOrEmpty(query))
        {
            return 0;
        }

        if (string.IsNullOrEmpty(target))
        {
            return null;
        }

        var queryIndex = 0;
        var score = 0;
        var consecutiveRun = 0;

        for (var targetIndex = 0; targetIndex < target.Length && queryIndex < query.Length; targetIndex++)
        {
            if (char.ToLowerInvariant(target[targetIndex]) == char.ToLowerInvariant(query[queryIndex]))
            {
                consecutiveRun++;
                score += consecutiveRun;

                if (targetIndex == 0)
                {
                    score += 1;
                }

                queryIndex++;
            }
            else
            {
                consecutiveRun = 0;
            }
        }

        return queryIndex == query.Length ? score : null;
    }
}

using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Tests.Commands;

public sealed class FuzzyMatcherTests
{
    [Fact]
    public void Score_EmptyQuery_MatchesWithZeroScore()
    {
        var score = FuzzyMatcher.Score(string.Empty, "Run Query");

        Assert.Equal(0, score);
    }

    [Fact]
    public void Score_ExactSubsequence_Matches()
    {
        var score = FuzzyMatcher.Score("run", "Run Query");

        Assert.NotNull(score);
    }

    [Fact]
    public void Score_CaseInsensitive_Matches()
    {
        var score = FuzzyMatcher.Score("RUN", "run query");

        Assert.NotNull(score);
    }

    [Fact]
    public void Score_OutOfOrderCharacters_DoesNotMatch()
    {
        var score = FuzzyMatcher.Score("qru", "Run Query");

        Assert.Null(score);
    }

    [Fact]
    public void Score_CharacterNotPresent_DoesNotMatch()
    {
        var score = FuzzyMatcher.Score("xyz", "Run Query");

        Assert.Null(score);
    }

    [Fact]
    public void Score_EmptyTarget_DoesNotMatchNonEmptyQuery()
    {
        var score = FuzzyMatcher.Score("run", string.Empty);

        Assert.Null(score);
    }

    [Fact]
    public void Score_ConsecutiveRun_ScoresHigherThanScattered()
    {
        var consecutive = FuzzyMatcher.Score("run", "Run Query");
        var scattered = FuzzyMatcher.Score("rqy", "Run Query");

        Assert.NotNull(consecutive);
        Assert.NotNull(scattered);
        Assert.True(consecutive > scattered);
    }
}

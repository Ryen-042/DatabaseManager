using System.Security;
using System.IO;
using System.Xml;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;

namespace DatabaseManager.Wpf.Editors;

internal static class SqlLanguageKeywords
{
    internal static readonly string[] Keywords =
    {
        "SELECT", "FROM", "WHERE", "GROUP BY", "ORDER BY", "HAVING", "JOIN", "INNER JOIN", "LEFT JOIN",
        "RIGHT JOIN", "FULL JOIN", "CROSS JOIN", "ON", "TOP", "DISTINCT", "INSERT INTO", "VALUES",
        "UPDATE", "SET", "DELETE", "TRUNCATE TABLE", "CREATE TABLE", "ALTER TABLE", "DROP TABLE",
        "CREATE PROCEDURE", "ALTER PROCEDURE", "DROP PROCEDURE", "EXEC", "BEGIN", "END", "CASE", "WHEN",
        "THEN", "ELSE", "DECLARE", "IF", "EXISTS", "IN", "BETWEEN", "LIKE", "UNION", "UNION ALL"
    };

    private static readonly Lazy<IReadOnlyList<string>> HighlightWords = new(CreateHighlightWords);

    internal static IReadOnlyList<string> GetHighlightWords()
    {
        return HighlightWords.Value;
    }

    private static IReadOnlyList<string> CreateHighlightWords()
    {
        return Keywords
            .SelectMany(static keyword => keyword.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}

internal static class SqlEditorSupport
{
    private static readonly Lazy<IHighlightingDefinition?> SqlHighlightingDefinition = new(CreateHighlightingDefinition);

    public static void Configure(TextEditor editor)
    {
        if (editor is null)
        {
            return;
        }

        editor.ShowLineNumbers = true;

        var highlightingDefinition = SqlHighlightingDefinition.Value;
        if (highlightingDefinition is not null)
        {
            editor.SyntaxHighlighting = highlightingDefinition;
        }

        var foldingManager = FoldingManager.Install(editor.TextArea);
        var foldingStrategy = new SqlFoldingStrategy();

        void UpdateFoldings()
        {
            if (editor.Document is null)
            {
                return;
            }

            try
            {
                foldingStrategy.UpdateFoldings(foldingManager, editor.Document);
            }
            catch
            {
                // Silently ignore folding errors to prevent UI crashes from text changes
            }
        }

        editor.TextChanged += (_, _) => UpdateFoldings();
        UpdateFoldings();
    }

    private static IHighlightingDefinition? CreateHighlightingDefinition()
    {
        try
        {
            var keywordWords = string.Join(
                Environment.NewLine,
                SqlLanguageKeywords.GetHighlightWords().Select(word => $"      <Word>{SecurityElement.Escape(word)}</Word>"));

            var xshd = $$"""
<?xml version="1.0"?>
<SyntaxDefinition name="SQL" extensions=".sql" xmlns="http://icsharpcode.net/sharpdevelop/syntaxhighlighting">
  <Color name="Comment" foreground="SeaGreen" />
  <Color name="String" foreground="SaddleBrown" />
  <Color name="Keyword" foreground="RoyalBlue" fontWeight="Bold" />
  <Color name="Number" foreground="DarkCyan" />
  <Color name="Parameter" foreground="DarkMagenta" />
  <RuleSet ignoreCase="true">
    <Span color="Comment" begin="--" end="$" multiline="true" />
    <Span color="Comment" begin="/\*" end="\*/" multiline="true" />
    <Rule color="String" regex="(?:N)?'(?:''|[^'])*'" />
    <Rule color="Number" regex="\b0x[0-9A-Fa-f]+\b" />
    <Rule color="Number" regex="\b\d+(?:\.\d+)?\b" />
    <Rule color="Parameter" regex="@@?[A-Za-z_][A-Za-z0-9_@$#]*" />
    <Rule color="Parameter" regex="\[[^\]]+\]" />
    <Keywords color="Keyword">
{{keywordWords}}
    </Keywords>
  </RuleSet>
</SyntaxDefinition>
""";

            using var reader = XmlReader.Create(new StringReader(xshd));
            return HighlightingLoader.Load(reader, new EmptyHighlightingDefinitionReferenceResolver());
        }
        catch
        {
            return null;
        }
    }

    private sealed class EmptyHighlightingDefinitionReferenceResolver : IHighlightingDefinitionReferenceResolver
    {
        public IHighlightingDefinition? GetDefinition(string name)
        {
            return null;
        }
    }
}

internal sealed class SqlFoldingStrategy
{
    public void UpdateFoldings(FoldingManager foldingManager, TextDocument document)
    {
        var foldings = CreateFoldings(document).ToList();
        foldingManager.UpdateFoldings(foldings, -1);
    }

    private static IEnumerable<NewFolding> CreateFoldings(TextDocument document)
    {
        var text = document.Text;
        if (string.IsNullOrEmpty(text))
        {
            yield break;
        }

        var blockStack = new Stack<int>();
        var index = 0;

        while (index < text.Length)
        {
            if (TryConsumeLineComment(text, ref index))
            {
                continue;
            }

            if (TryConsumeBlockComment(text, ref index, out var commentStart, out var commentEnd))
            {
                AddFoldingIfMultiline(document, commentStart + 2, commentEnd - 2, "/* ... */", out var folding);
                if (folding is not null)
                {
                    yield return folding;
                }

                continue;
            }

            if (TryConsumeStringLiteral(text, ref index))
            {
                continue;
            }

            if (TryConsumeBracketedIdentifier(text, ref index))
            {
                continue;
            }

            if (index >= text.Length)
            {
                break;
            }

            var current = text[index];
            if (TryConsumeWord(text, ref index, out var word, out var wordStart, out var wordEnd))
            {
                if (string.Equals(word, "BEGIN", StringComparison.OrdinalIgnoreCase))
                {
                    blockStack.Push(wordEnd);
                }
                else if (string.Equals(word, "END", StringComparison.OrdinalIgnoreCase) && blockStack.Count > 0)
                {
                    var start = blockStack.Pop();
                    AddFoldingIfMultiline(document, start, wordStart, "BEGIN ... END", out var folding);
                    if (folding is not null)
                    {
                        yield return folding;
                    }
                }

                continue;
            }

            index++;
        }
    }

    private static bool TryConsumeLineComment(string text, ref int index)
    {
        if (index + 1 >= text.Length || text[index] != '-' || text[index + 1] != '-')
        {
            return false;
        }

        index += 2;
        while (index < text.Length && text[index] is not '\r' and not '\n')
        {
            index++;
        }

        return true;
    }

    private static bool TryConsumeBlockComment(string text, ref int index, out int start, out int end)
    {
        start = index;
        end = index;

        if (index + 1 >= text.Length || text[index] != '/' || text[index + 1] != '*')
        {
            return false;
        }

        index += 2;
        while (index + 1 < text.Length)
        {
            if (text[index] == '*' && text[index + 1] == '/')
            {
                index += 2;
                end = index;
                return true;
            }

            index++;
        }

        end = text.Length;
        return true;
    }

    private static bool TryConsumeStringLiteral(string text, ref int index)
    {
        if (index >= text.Length || text[index] != '\'')
        {
            return false;
        }

        index++;
        while (index < text.Length)
        {
            if (text[index] == '\'' && index + 1 < text.Length && text[index + 1] == '\'')
            {
                index += 2;
                continue;
            }

            if (text[index] == '\'')
            {
                index++;
                return true;
            }

            index++;
        }

        return true;
    }

    private static bool TryConsumeBracketedIdentifier(string text, ref int index)
    {
        if (index >= text.Length || text[index] != '[')
        {
            return false;
        }

        index++;
        while (index < text.Length)
        {
            if (text[index] == ']')
            {
                index++;
                return true;
            }

            index++;
        }

        return true;
    }

    private static bool TryConsumeWord(string text, ref int index, out string word, out int wordStart, out int wordEnd)
    {
        word = string.Empty;
        wordStart = index;
        wordEnd = index;

        if (index >= text.Length || (!char.IsLetter(text[index]) && text[index] != '_'))
        {
            return false;
        }

        wordStart = index;
        index++;
        while (index < text.Length && (char.IsLetterOrDigit(text[index]) || text[index] is '_' or '$'))
        {
            index++;
        }

        wordEnd = index;
        word = text[wordStart..wordEnd];
        return true;
    }

    private static void AddFoldingIfMultiline(TextDocument document, int startOffset, int endOffset, string title, out NewFolding? folding)
    {
        folding = null;

        if (document?.Text == null || startOffset < 0 || endOffset <= startOffset || endOffset > document.TextLength)
        {
            return;
        }

        try
        {
            var content = document.Text[startOffset..endOffset];
            if (!content.Contains('\n') && !content.Contains('\r'))
            {
                return;
            }

            folding = new NewFolding(startOffset, endOffset)
            {
                Name = title,
                DefaultClosed = false
            };
        }
        catch
        {
            // Silently ignore any indexing errors
        }
    }
}
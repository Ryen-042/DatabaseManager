using System.Windows;
using ICSharpCode.AvalonEdit;

namespace DatabaseManager.Wpf.Editors;

public sealed class AvalonEditSqlTextEditorAdapter : ISqlTextEditor
{
    private readonly TextEditor _editor;

    public AvalonEditSqlTextEditorAdapter(TextEditor editor)
    {
        _editor = editor;
    }

    public string Text
    {
        get => _editor.Text;
        set => _editor.Text = value;
    }

    public int CaretIndex
    {
        get => _editor.CaretOffset;
        set => _editor.CaretOffset = Math.Clamp(value, 0, _editor.Text.Length);
    }

    public UIElement Element => _editor;

    public void Select(int start, int length)
    {
        var safeStart = Math.Clamp(start, 0, _editor.Text.Length);
        var safeLength = Math.Clamp(length, 0, _editor.Text.Length - safeStart);
        _editor.Select(safeStart, safeLength);
    }

    public void ReplaceSelection(string text)
    {
        var selectionStart = _editor.SelectionStart;
        _editor.Document.Replace(selectionStart, _editor.SelectionLength, text);
        _editor.CaretOffset = selectionStart + text.Length;
        _editor.Select(_editor.CaretOffset, 0);
    }

    public Rect GetCaretRect()
    {
        return _editor.TextArea.Caret.CalculateCaretRectangle();
    }

    public void Focus()
    {
        _editor.Focus();
    }
}

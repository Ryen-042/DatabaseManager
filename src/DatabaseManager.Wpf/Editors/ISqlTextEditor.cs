using System.Windows;

namespace DatabaseManager.Wpf.Editors;

public interface ISqlTextEditor
{
    string Text { get; set; }

    int CaretIndex { get; set; }

    UIElement Element { get; }

    void Select(int start, int length);

    void ReplaceSelection(string text);

    Rect GetCaretRect();

    void Focus();
}

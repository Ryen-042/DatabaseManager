using System.Windows.Input;

namespace DatabaseManager.Wpf.Commands;

public static class KeyGestureFormatter
{
    public static string Format(KeyGesture? gesture)
    {
        if (gesture is null)
        {
            return string.Empty;
        }

        var parts = new List<string>();
        if (gesture.Modifiers.HasFlag(ModifierKeys.Control))
        {
            parts.Add("Ctrl");
        }

        if (gesture.Modifiers.HasFlag(ModifierKeys.Alt))
        {
            parts.Add("Alt");
        }

        if (gesture.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            parts.Add("Shift");
        }

        parts.Add(gesture.Key.ToString());
        return string.Join("+", parts);
    }
}

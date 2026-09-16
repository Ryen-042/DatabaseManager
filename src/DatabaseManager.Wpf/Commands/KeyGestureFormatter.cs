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

        parts.Add(FormatKey(gesture.Key));
        return string.Join("+", parts);
    }

    /// <summary>
    /// WPF's <see cref="Key"/> enum names digits "D0".."D9" (can't start an identifier with a
    /// digit) and gives OEM/punctuation keys internal names that don't match the character
    /// they actually produce on a US keyboard - translate the ones this app's shortcuts
    /// actually use back to what a user would expect to read, falling back to the raw enum
    /// name for anything else.
    /// </summary>
    private static string FormatKey(Key key) => key switch
    {
        >= Key.D0 and <= Key.D9 => ((int)(key - Key.D0)).ToString(),
        >= Key.NumPad0 and <= Key.NumPad9 => ((int)(key - Key.NumPad0)).ToString(),
        Key.OemQuestion => "/",
        Key.OemComma => ",",
        Key.OemPeriod => ".",
        Key.OemMinus => "-",
        Key.OemPlus => "=",
        Key.OemSemicolon => ";",
        Key.OemQuotes => "'",
        Key.OemBackslash => "\\",
        Key.OemOpenBrackets => "[",
        Key.OemCloseBrackets => "]",
        Key.OemTilde => "`",
        Key.Space => "Space",
        _ => key.ToString()
    };
}

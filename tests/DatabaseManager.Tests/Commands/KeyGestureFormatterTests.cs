using System.Windows.Input;
using DatabaseManager.Wpf.Commands;

namespace DatabaseManager.Tests.Commands;

public sealed class KeyGestureFormatterTests
{
    [Fact]
    public void Format_NullGesture_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, KeyGestureFormatter.Format(null));
    }

    [Theory]
    [InlineData(Key.D1, "Ctrl+1")]
    [InlineData(Key.D5, "Ctrl+5")]
    [InlineData(Key.D0, "Ctrl+0")]
    public void Format_DigitKeys_ShowsPlainDigitNotDEnumName(Key key, string expected)
    {
        var gesture = new KeyGesture(key, ModifierKeys.Control);

        Assert.Equal(expected, KeyGestureFormatter.Format(gesture));
    }

    [Fact]
    public void Format_OemQuestion_ShowsSlash()
    {
        var gesture = new KeyGesture(Key.OemQuestion, ModifierKeys.Control);

        Assert.Equal("Ctrl+/", KeyGestureFormatter.Format(gesture));
    }

    [Fact]
    public void Format_MultipleModifiers_OrdersControlAltShift()
    {
        var gesture = new KeyGesture(Key.P, ModifierKeys.Control | ModifierKeys.Shift);

        Assert.Equal("Ctrl+Shift+P", KeyGestureFormatter.Format(gesture));
    }

    [Fact]
    public void Format_PlainLetterKey_UsesEnumName()
    {
        var gesture = new KeyGesture(Key.R, ModifierKeys.Control);

        Assert.Equal("Ctrl+R", KeyGestureFormatter.Format(gesture));
    }
}

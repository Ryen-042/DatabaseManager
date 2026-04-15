using System.Collections;
using System.Data.SqlTypes;

namespace DatabaseManager.Core.Services;

public static class DisplayValueFormatter
{
    private const int BinaryPreviewBytes = 64;
    private const int ArrayPreviewItems = 20;

    public static string FormatForDisplay(object? value, bool fullOutput = false)
    {
        if (value is null or DBNull)
        {
            return string.Empty;
        }

        return value switch
        {
            byte[] bytes => FormatBinary(bytes, fullOutput),
            SqlBinary sqlBinary => sqlBinary.IsNull ? string.Empty : FormatBinary(sqlBinary.Value, fullOutput),
            SqlBytes sqlBytes => sqlBytes.IsNull ? string.Empty : FormatBinary(sqlBytes.Value, fullOutput),
            char[] chars => new string(chars),
            Array array => FormatArray(array, fullOutput),
            _ => value.ToString() ?? string.Empty
        };
    }

    private static string FormatBinary(ReadOnlySpan<byte> bytes, bool fullOutput)
    {
        if (bytes.Length == 0)
        {
            return "0x";
        }

        if (fullOutput || bytes.Length <= BinaryPreviewBytes)
        {
            return $"0x{Convert.ToHexString(bytes)}";
        }

        var preview = Convert.ToHexString(bytes[..BinaryPreviewBytes]);
        return $"0x{preview}... ({bytes.Length} bytes)";
    }

    private static string FormatArray(Array array, bool fullOutput)
    {
        if (array.Length == 0)
        {
            return "[]";
        }

        var limit = fullOutput ? array.Length : ArrayPreviewItems;

        var values = array.Cast<object?>()
            .Take(limit)
            .Select(item => item?.ToString() ?? "NULL")
            .ToArray();

        var suffix = !fullOutput && array.Length > ArrayPreviewItems ? ", ..." : string.Empty;
        return $"[{string.Join(", ", values)}{suffix}]";
    }
}
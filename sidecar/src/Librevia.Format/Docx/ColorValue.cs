using System.Globalization;

namespace Librevia.Format.Docx;

/// <summary>
/// CSS color → six-digit hex, OOXML's only form. Alpha is dropped: OOXML has no transparent text. A
/// color that does not convert returns <c>null</c>, and the caller records the loss.
/// </summary>
internal static class ColorValue
{
    /// <summary>
    /// The 16 HTML colors, the greys and Word's theme colors; the rest of CSS's 148 go to the
    /// warning.
    /// </summary>
    private static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000",
        ["silver"] = "C0C0C0",
        ["gray"] = "808080",
        ["grey"] = "808080",
        ["white"] = "FFFFFF",
        ["maroon"] = "800000",
        ["red"] = "FF0000",
        ["purple"] = "800080",
        ["fuchsia"] = "FF00FF",
        ["magenta"] = "FF00FF",
        ["green"] = "008000",
        ["lime"] = "00FF00",
        ["olive"] = "808000",
        ["yellow"] = "FFFF00",
        ["navy"] = "000080",
        ["blue"] = "0000FF",
        ["teal"] = "008080",
        ["aqua"] = "00FFFF",
        ["cyan"] = "00FFFF",
        ["orange"] = "FFA500",
        ["pink"] = "FFC0CB",
        ["brown"] = "A52A2A",
        ["gold"] = "FFD700",
        ["darkgray"] = "A9A9A9",
        ["darkgrey"] = "A9A9A9",
        ["lightgray"] = "D3D3D3",
        ["lightgrey"] = "D3D3D3",
        ["whitesmoke"] = "F5F5F5",
    };

    /// <summary>Six-digit hex, or <c>null</c> if it cannot be done.</summary>
    public static string? Hex(string? css)
    {
        if (string.IsNullOrWhiteSpace(css)) return null;

        var value = css.Trim();

        // No color: black would add color where there is none.
        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("currentcolor", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("transparent", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (Named.TryGetValue(value, out var named)) return named;
        if (value.StartsWith('#')) return FromHex(value[1..]);
        if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) return FromRgb(value);

        // Only OOXML's six digits: in short form, `fade` would become FFAADD.
        return value.Length == 6 ? FromHex(value) : null;
    }

    private static string? FromHex(string digits)
    {
        if (!digits.All(Uri.IsHexDigit)) return null;

        // The alpha of `#rrggbbaa` and `#rgba` is dropped.
        return digits.Length switch
        {
            3 or 4 => string.Concat(digits[..3].Select(digit => new string(digit, 2))).ToUpperInvariant(),
            6 or 8 => digits[..6].ToUpperInvariant(),
            _ => null,
        };
    }

    /// <summary>
    /// `rgb(255, 0, 0)` and `rgba(255 0 0 / 50%)`, which is what the browser returns.
    /// </summary>
    private static string? FromRgb(string value)
    {
        var open = value.IndexOf('(', StringComparison.Ordinal);
        var close = value.LastIndexOf(')');
        if (open < 0 || close < open) return null;

        var parts = value[(open + 1)..close]
            .Split([',', ' ', '/'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 3) return null;

        var channels = new int[3];
        for (var index = 0; index < 3; index++)
        {
            if (Channel(parts[index]) is not { } channel) return null;
            channels[index] = channel;
        }

        return $"{channels[0]:X2}{channels[1]:X2}{channels[2]:X2}";
    }

    private static int? Channel(string text)
    {
        var percent = text.EndsWith('%');
        var digits = percent ? text[..^1] : text;
        if (!double.TryParse(digits, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return null;

        var scaled = percent ? number * 255 / 100 : number;
        return (int)Math.Clamp(Math.Round(scaled), 0, 255);
    }
}

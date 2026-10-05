namespace Librevia.Format.Docx;

/// <summary>
/// The natural line height, in multiples of the font size: <c>(ascender - descender + lineGap) /
/// unitsPerEm</c> from the <c>hhea</c> table, the math of Word, LibreOffice and <c>line-height:
/// normal</c>. The OOXML multiple is relative to it, and stating it as a number keeps Chromium from
/// rounding it to a whole pixel. Only the fonts the installer ships and the ones they substitute:
/// for others, the substitute depends on the machine.
/// </summary>
internal static class LineMetrics
{
    /// <summary>The editor font when the document names none.</summary>
    private const double LiberationSerif = 1.1499;

    private static readonly Dictionary<string, double> Known = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Arial"] = 1.1499,
        ["Helvetica"] = 1.1499,
        ["Liberation Sans"] = 1.1499,
        ["Times New Roman"] = LiberationSerif,
        ["Liberation Serif"] = LiberationSerif,
        ["Courier New"] = 1.1328,
        ["Liberation Mono"] = 1.1328,
        ["Calibri"] = 1.2207,
        ["Carlito"] = 1.2207,
        ["Cambria"] = 1.15,
        ["Caladea"] = 1.15,
    };

    /// <summary>Null when it is unknown which font file the browser will use.</summary>
    public static double? Of(string? font) =>
        string.IsNullOrWhiteSpace(font)
            ? LiberationSerif
            : Known.TryGetValue(font.Trim(), out var natural) ? natural : null;
}

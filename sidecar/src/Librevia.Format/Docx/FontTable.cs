using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Each font's generic family (<c>w:family</c> in <c>word/fontTable.xml</c>), so the substitute for
/// a missing font, like Segoe UI or Aptos, does not fall into the serif at the end of the stack.
/// </summary>
public sealed class FontTable
{
    private readonly Dictionary<string, string> _generic = new(StringComparer.OrdinalIgnoreCase);

    public FontTable(MainDocumentPart part)
    {
        var fonts = part.FontTablePart?.Fonts;
        if (fonts is null) return;

        foreach (var font in fonts.Elements<Font>())
        {
            if (font.Name?.Value is not { Length: > 0 } name) continue;
            if (GenericOf(font.FontFamily?.Val?.Value) is { } generic) _generic[name] = generic;
        }
    }

    /// <summary>The requested font and the generic substitute, for CSS.</summary>
    public string Stack(string name)
    {
        var trimmed = name.Trim();
        return _generic.TryGetValue(trimmed, out var generic) ? $"{trimmed}, {generic}" : trimmed;
    }

    private static string? GenericOf(FontFamilyValues? family)
    {
        if (family is null) return null;
        if (family == FontFamilyValues.Swiss) return "sans-serif";
        if (family == FontFamilyValues.Roman) return "serif";
        if (family == FontFamilyValues.Modern) return "monospace";
        if (family == FontFamilyValues.Script) return "cursive";
        if (family == FontFamilyValues.Decorative) return "fantasy";
        return null;
    }
}

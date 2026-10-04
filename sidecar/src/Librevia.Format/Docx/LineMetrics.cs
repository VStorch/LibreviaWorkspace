namespace Librevia.Format.Docx;

/// <summary>
/// A altura natural da linha, em múltiplos do tamanho da fonte:
/// <c>(ascender - descender + lineGap) / unitsPerEm</c> da tabela <c>hhea</c>, a
/// conta do Word, do LibreOffice e do <c>line-height: normal</c>. O múltiplo do
/// OOXML é sobre ela, e dizê-la em número evita que o Chromium a arredonde para
/// pixel inteiro. Só as fontes que o instalador leva e as que elas substituem: para
/// as outras, a substituta depende da máquina.
/// </summary>
internal static class LineMetrics
{
    /// <summary>A fonte do editor quando o documento não diz outra.</summary>
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

    /// <summary>Nula quando não se sabe que arquivo de fonte o navegador vai usar.</summary>
    public static double? Of(string? font) =>
        string.IsNullOrWhiteSpace(font)
            ? LiberationSerif
            : Known.TryGetValue(font.Trim(), out var natural) ? natural : null;
}

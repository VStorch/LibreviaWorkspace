namespace Librevia.Format.Docx;

/// <summary>
/// Altura natural da linha de cada fonte, em múltiplos do tamanho dela.
/// </summary>
/// <remarks>
/// É `(ascender - descender + lineGap) / unitsPerEm` lido da tabela `hhea`, que
/// é a conta que o Word e o LibreOffice fazem para saber quanto ocupa uma linha
/// "simples" — e a mesma que o navegador chama de `line-height: normal`.
///
/// Sabê-la aqui resolve duas coisas de uma vez:
///
/// <list type="number">
/// <item>O múltiplo do OOXML é sobre **essa** altura, e não sobre o tamanho da
/// fonte. "1,13 linha" em Arial 10 pt são 12,98 pt no LibreOffice, e eram 11,3
/// aqui — quase uma linha a mais por página.</item>
/// <item>O Chromium arredonda as métricas para pixel inteiro ao resolver
/// `normal`: Arial 10 pt dava 15 px em vez de 15,33. São 2 % por linha, o
/// bastante para um documento de quinze folhas fechar em dezesseis. Dizer a
/// altura em número evita o arredondamento.</item>
/// </list>
///
/// Só as fontes que o instalador leva, e as que elas substituem. Para o resto
/// não há palpite honesto — a substituta depende da máquina — e o leitor
/// continua deixando a decisão com o navegador.
/// </remarks>
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

    /// <summary>
    /// A altura natural da fonte, ou <c>null</c> quando não se sabe qual arquivo
    /// o navegador vai usar.
    /// </summary>
    public static double? Of(string? font) =>
        string.IsNullOrWhiteSpace(font)
            ? LiberationSerif
            : Known.TryGetValue(font.Trim(), out var natural) ? natural : null;
}

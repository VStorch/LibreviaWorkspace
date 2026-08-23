using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A que família genérica cada fonte do documento pertence.
/// </summary>
/// <remarks>
/// Um documento do Word nomeia fontes que a máquina de quem abre pode não ter —
/// Segoe UI, Aptos, Calibri. Quando a fonte falta, o navegador cai na próxima da
/// pilha, e a pilha do editor termina em serifa: a capa do modelo de manual, que
/// pede Segoe UI, saía com o título em Times enquanto o LibreOffice o desenha
/// sem serifa.
///
/// O arquivo diz de que tipo cada fonte é, em `word/fontTable.xml`
/// (`w:family`), e é essa a informação que falta para escolher a substituta
/// certa. Não é a mesma coisa que a fonte metricamente compatível que o
/// instalador leva — essa acerta a medida da linha, e vale para as poucas
/// famílias que empacotamos. Esta aqui é a rede embaixo, para todo o resto.
/// </remarks>
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

    /// <summary>
    /// A fonte pedida seguida da substituta genérica, prontas para o CSS.
    /// </summary>
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
        // `auto` não diz nada: sem informação, é melhor não inventar uma pilha.
        return null;
    }
}

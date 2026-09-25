using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A definição de uma lista como o editor a leva: nove níveis, cada um com
/// formato, texto, início e recuo.
/// </summary>
/// <remarks>
/// É o espelho de `list-numbering.ts` (`NumberingDef`/`LevelDef`), e os dois
/// precisam mudar juntos. Mora no **nó da lista**, e não num catálogo à parte
/// como os estilos, por uma razão prática: o nó viaja sozinho — colado de outro
/// documento, gravado no `.sdoc`, desfeito e refeito — e a numeração tem de ir
/// com ele. O custo é que a definição entra na impressão digital do item de fora
/// quando a lista é aninhada; por isso leitor e editor a produzem idêntica, e o
/// editor nunca a reescreve sozinho.
///
/// A marca da lista com marcador vai **traduzida** (`•`, e não o `` da fonte
/// Symbol): é o que a tela desenha. A volta ao glifo e à fonte que o Word espera
/// é de <see cref="GlyphOf"/>, e as duas tabelas são a mesma lida nos dois
/// sentidos.
/// </remarks>
public static class ListLevels
{
    public const int Count = 9;

    /// <summary>Recuo por nível e marcador pendurado: as medidas do próprio Word.</summary>
    public const int IndentStepTwips = 720;
    public const int HangingTwips = 360;

    /// <summary>Glifo da área de uso privado, com a fonte que o desenha, e a marca que a tela usa.</summary>
    private static readonly (char Glyph, string? Font, char Shown)[] Glyphs =
    [
        ('', "Symbol", '•'),
        ('o', "Courier New", 'o'),
        ('', "Wingdings", '▪'),
        ('', "Wingdings", '➢'),
        ('', "Wingdings", '●'),
        ('', "Wingdings", '□'),
        ('', "Wingdings", '✓'),
    ];

    /// <summary>
    /// A marca de um `w:lvlText` de marcador, como qualquer fonte a desenha.
    /// </summary>
    /// <remarks>
    /// O Word grava os glifos das fontes Symbol e Wingdings na área de uso
    /// privado do Unicode. Servido como está, aparece a caixinha de caractere
    /// ausente; trocado pelo equivalente de verdade, aparece a marca que o
    /// documento pede — e não a bolinha que o CSS escolhe sozinho. O que a
    /// tabela não conhece vira bolinha: é o que há de mais neutro.
    /// </remarks>
    public static string Shown(string text) => string.Concat(text.Select(letter =>
    {
        foreach (var (glyph, font, shown) in Glyphs)
        {
            if (glyph == letter && font is not null && letter >= '') return shown;
        }

        return letter switch
        {
            '' => '▪',
            >= '' and <= '' => '•',
            _ => letter,
        };
    }));

    /// <summary>O glifo e a fonte que o Word espera para uma marca da tela.</summary>
    public static (string Text, string? Font) GlyphOf(string shown)
    {
        if (shown.Length == 1)
        {
            foreach (var (glyph, font, mark) in Glyphs)
            {
                if (mark == shown[0]) return (glyph.ToString(), font);
            }
        }

        return (shown, null);
    }

    /// <summary>
    /// Os níveis que o Word dá a uma lista nova: 1. a. i. na numerada, e • o ▪
    /// na com marcador, repetidos de três em três.
    /// </summary>
    /// <remarks>Iguais a `defaultLevels` em `list-numbering.ts`: a tela desenha a lista nova com eles.</remarks>
    public static JsonArray Defaults(string kind)
    {
        var bullet = string.Equals(kind, "bulletList", StringComparison.Ordinal);
        string[] formats = ["decimal", "lowerLetter", "lowerRoman"];
        string[] marks = ["•", "o", "▪"];

        var levels = new JsonArray();
        for (var level = 0; level < Count; level++)
        {
            levels.Add(Level(
                bullet ? "bullet" : formats[level % 3],
                bullet ? marks[level % 3] : $"%{level + 1}.",
                1,
                Mm((level + 1) * IndentStepTwips),
                Mm(HangingTwips)));
        }

        return levels;
    }

    public static JsonObject Level(string format, string text, int start, double? indentMm, double? hangingMm, bool legal = false)
    {
        var level = new JsonObject
        {
            ["fmt"] = format,
            ["text"] = text,
            ["start"] = start,
        };
        if (indentMm is { } indent) level["indentMm"] = indent;
        if (hangingMm is { } hanging) level["hangingMm"] = hanging;
        if (legal) level["legal"] = true;
        return level;
    }

    public static double Mm(int twips) => Math.Round(twips / 1440.0 * 25.4, 2);

    /// <summary>O nome OOXML do formato, como o `w:numFmt/@w:val` o grava.</summary>
    public static string FormatName(Level? definition) =>
        definition?.NumberingFormat?.Val?.InnerText is { Length: > 0 } name ? name : "decimal";

    /// <summary>
    /// Um nível do OOXML, montado de volta da definição do editor.
    /// </summary>
    /// <remarks>
    /// Formato que o SDK não conhece vira decimal: gravar um valor fora da
    /// enumeração faria o Word recusar o arquivo inteiro.
    /// </remarks>
    public static Level ToOpenXml(JsonObject? source, int index, string kind)
    {
        var fallback = (JsonObject)Defaults(kind)[index]!;
        var level = source ?? fallback;

        var format = level["fmt"]?.GetValue<string>() ?? fallback["fmt"]!.GetValue<string>();
        var text = level["text"]?.GetValue<string>() ?? string.Empty;
        var start = level["start"]?.GetValue<int>() ?? 1;
        var indent = Number(level["indentMm"]) ?? Number(fallback["indentMm"])!.Value;
        var hanging = Number(level["hangingMm"]) ?? 0;

        var bullet = format == "bullet";
        var (glyph, font) = bullet ? GlyphOf(text) : (text, null);

        var definition = new Level
        {
            LevelIndex = index,
            StartNumberingValue = new StartNumberingValue { Val = start },
            NumberingFormat = new NumberingFormat { Val = FormatOf(format) },
        };
        if (level["legal"]?.GetValue<bool>() == true) definition.IsLegalNumberingStyle = new IsLegalNumberingStyle();
        definition.LevelText = new LevelText { Val = glyph };
        definition.LevelJustification = new LevelJustification { Val = LevelJustificationValues.Left };
        definition.PreviousParagraphProperties = new PreviousParagraphProperties(new Indentation
        {
            Left = AttrMm(indent),
            Hanging = AttrMm(hanging),
        });

        if (font is not null)
        {
            // A marca do Word vem da área de uso privado, e só a fonte dela a
            // desenha. Sem declarar a fonte aparece a caixinha de caractere ausente.
            definition.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(new RunFonts
            {
                Ascii = font,
                HighAnsi = font,
                Hint = FontTypeHintValues.Default,
            });
        }

        return definition;
    }

    private static string AttrMm(double mm) =>
        Attr.MmToTwips(mm).ToString(CultureInfo.InvariantCulture);

    private static double? Number(JsonNode? node) =>
        node is JsonValue value && value.TryGetValue<double>(out var number) ? number : null;

    private static NumberFormatValues FormatOf(string name) => name switch
    {
        "bullet" => NumberFormatValues.Bullet,
        "none" => NumberFormatValues.None,
        "decimalZero" => NumberFormatValues.DecimalZero,
        "lowerLetter" => NumberFormatValues.LowerLetter,
        "upperLetter" => NumberFormatValues.UpperLetter,
        "lowerRoman" => NumberFormatValues.LowerRoman,
        "upperRoman" => NumberFormatValues.UpperRoman,
        _ => NumberFormatValues.Decimal,
    };
}

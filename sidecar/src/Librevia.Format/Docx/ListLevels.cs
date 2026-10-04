using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A lista como o editor a leva: nove níveis com formato, texto, início e recuo.
/// Espelho de <c>NumberingDef</c>/<c>LevelDef</c> em <c>list-numbering.ts</c>: mudam
/// juntos. Mora no nó para viajar com ele (colar, <c>.sdoc</c>, desfazer), e por
/// isso entra na impressão digital: leitor e editor a produzem idêntica.
/// O marcador vai traduzido (<c>•</c>, e não o glifo da Symbol); <see cref="GlyphOf"/>
/// faz a volta.
/// </summary>
public static class ListLevels
{
    public const int Count = 9;

    /// <summary>Recuo por nível e marcador pendurado: as medidas do próprio Word.</summary>
    public const int IndentStepTwips = Unit.IndentStepTwips;
    public const int HangingTwips = IndentStepTwips / 2;

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
    /// O Word grava os marcadores da Symbol e da Wingdings na área de uso privado,
    /// que fora dessas fontes é a caixinha de caractere ausente. O desconhecido vira bolinha.
    /// </summary>
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
    /// Os níveis do Word para lista nova (1. a. i. e • o ▪), iguais a
    /// <c>defaultLevels</c> em <c>list-numbering.ts</c>.
    /// </summary>
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

    public static double Mm(int twips) => Math.Round(twips / (double)Unit.TwipsPerInch * Unit.MillimetersPerInch, 2);

    /// <summary>O nome OOXML do formato, como o `w:numFmt/@w:val` o grava.</summary>
    public static string FormatName(Level? definition) =>
        definition?.NumberingFormat?.Val?.InnerText is { Length: > 0 } name ? name : "decimal";

    /// <summary>Formato que o SDK não conhece vira decimal: o Word recusaria o arquivo.</summary>
    public static Level ToOpenXml(JsonObject? source, int index, string kind, Inventory? inventory = null)
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
            NumberingFormat = new NumberingFormat { Val = FormatOf(format, inventory) },
        };
        if (level["restart"] is JsonValue restart && restart.TryGetValue<int>(out var restartValue))
        {
            definition.LevelRestart = new LevelRestart { Val = restartValue };
        }
        if (level["legal"]?.GetValue<bool>() == true) definition.IsLegalNumberingStyle = new IsLegalNumberingStyle();
        if (level["suff"]?.GetValue<string>() is { } suffix)
        {
            definition.LevelSuffix = new LevelSuffix
            {
                Val = suffix == "space" ? LevelSuffixValues.Space : suffix == "nothing" ? LevelSuffixValues.Nothing : LevelSuffixValues.Tab,
            };
        }
        definition.LevelText = new LevelText { Val = glyph };
        definition.LevelJustification = new LevelJustification
        {
            Val = level["jc"]?.GetValue<string>() switch
            {
                "right" or "end" => LevelJustificationValues.Right,
                "center" => LevelJustificationValues.Center,
                _ => LevelJustificationValues.Left,
            },
        };

        // Sem o original, o formato do número e o estilo ligado ficaram no documento de origem.
        if (level["extra"]?.GetValue<bool>() == true)
        {
            inventory?.NoteLoss("formatação própria do número de uma lista (fonte, cor ou estilo do nível)");
        }
        definition.PreviousParagraphProperties = new PreviousParagraphProperties(new Indentation
        {
            Left = AttrMm(indent),
            Hanging = AttrMm(hanging),
        });

        if (font is not null)
        {
            // O glifo da área de uso privado só a fonte dele desenha.
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

    /// <summary>
    /// Qualquer nome do esquema, não só os que a tela desenha. Fora dele, decimal com
    /// aviso: o Word recusaria o arquivo.
    /// </summary>
    private static NumberFormatValues FormatOf(string name, Inventory? inventory)
    {
        var value = new NumberFormatValues(name);
        if (((DocumentFormat.OpenXml.IEnumValue)value).IsValid) return value;

        inventory?.NoteLoss($"formato de numeração de lista \"{name}\" (gravado como decimal)");
        return NumberFormatValues.Decimal;
    }
}

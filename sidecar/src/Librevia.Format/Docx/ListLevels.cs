using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A list as the editor carries it: nine levels with format, text, start and indent. Mirrors
/// <c>NumberingDef</c>/<c>LevelDef</c> in <c>list-numbering.ts</c>: they change together. It lives
/// on the node to travel with it (paste, <c>.sdoc</c>, undo), so it enters the fingerprint: reader
/// and editor produce it identically. The marker goes translated (<c>•</c>, not the Symbol glyph);
/// <see cref="GlyphOf"/> converts back.
/// </summary>
public static class ListLevels
{
    public const int Count = 9;

    /// <summary>Indent per level and hanging marker: Word's own measures.</summary>
    public const int IndentStepTwips = Unit.IndentStepTwips;
    public const int HangingTwips = IndentStepTwips / 2;

    /// <summary>
    /// A private use area glyph, the font that draws it, and the mark the screen uses.
    /// </summary>
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
    /// Word writes Symbol and Wingdings bullets in the private use area, which outside those fonts
    /// is the missing-character box. An unknown one becomes a dot.
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

    /// <summary>The glyph and font Word expects for a screen mark.</summary>
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
    /// Word's levels for a new list (1. a. i. and • o ▪), the same as <c>defaultLevels</c> in
    /// <c>list-numbering.ts</c>.
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

    /// <summary>The OOXML format name, as `w:numFmt/@w:val` writes it.</summary>
    public static string FormatName(Level? definition) =>
        definition?.NumberingFormat?.Val?.InnerText is { Length: > 0 } name ? name : "decimal";

    /// <summary>
    /// A format the SDK does not know becomes decimal: Word would refuse the file.
    /// </summary>
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

        // Without the original, the number format and the linked style stayed in the source
        // document.
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
            // Only its own font draws a private use area glyph.
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
    /// Any schema name, not only those the screen draws. Outside it, decimal with a warning: Word
    /// would refuse the file.
    /// </summary>
    private static NumberFormatValues FormatOf(string name, Inventory? inventory)
    {
        var value = new NumberFormatValues(name);
        if (((DocumentFormat.OpenXml.IEnumValue)value).IsValid) return value;

        inventory?.NoteLoss($"formato de numeração de lista \"{name}\" (gravado como decimal)");
        return NumberFormatValues.Decimal;
    }
}

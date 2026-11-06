using System.Globalization;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>office:styles</c> → the editor's style sheet: the paragraph and text families. The id is the
/// file's <c>style:name</c> (what paragraphs point to); the name, its <c>style:display-name</c>.
/// LibreOffice's built-ins take Word's id and name instead (<see cref="OdtWordStyles"/>).
/// </summary>
internal static partial class OdtStyleSheet
{
    private const int WordOutlineLevels = 9;

    public static StyleSheetDto Read(OdtStyles styles)
    {
        var sheet = new Dictionary<string, StyleDefinitionDto>(StringComparer.Ordinal);
        foreach (var style in styles.Named)
        {
            if (Definition(styles, style) is { } definition) sheet.TryAdd(definition.Id, definition);
        }

        var standard = styles.IsNamed(OdtStyles.ParagraphFamily, OdtStyles.DefaultParagraphStyle)
            ? OdtWordStyles.IdOf(OdtStyles.ParagraphFamily, OdtStyles.DefaultParagraphStyle)
            : null;
        var defaults = new StyleDefaultsDto(
            styles.DefaultParagraph().ToDto(outlineLevel: null),
            styles.DefaultCharacter().ToDto(),
            standard,
            CharacterStyleId: null);
        return new StyleSheetDto(defaults, sheet);
    }

    private static StyleDefinitionDto? Definition(OdtStyles styles, OdtStyle style)
    {
        var paragraph = style.Family == OdtStyles.ParagraphFamily;
        if (!paragraph && style.Family != OdtStyles.TextFamily) return null;

        var parent = style.Parent is { } named && styles.IsNamed(style.Family, named)
            ? OdtWordStyles.IdOf(style.Family, named)
            : null;
        var next = paragraph && style.Next is { } following && styles.IsNamed(style.Family, following)
            ? OdtWordStyles.IdOf(style.Family, following)
            : null;
        return new StyleDefinitionDto(
            Id: style.Id,
            Name: OdtWordStyles.NameOf(style.Family, style.Name) ?? style.DisplayName ?? DisplayNameOf(style.Name),
            Type: paragraph ? "paragraph" : "character",
            QFormat: true,
            Hidden: false,
            Custom: false,
            BasedOn: parent,
            Next: next,
            Paragraph: paragraph ? ParagraphOf(style) : null,
            Character: CharacterOf(styles, style));
    }

    private static StyleParagraphDto? ParagraphOf(OdtStyle style)
    {
        var dto = ParagraphFormat.Of(style.Properties("paragraph-properties")).ToDto(WordOutlineLevelOf(style.OutlineLevel));
        return dto == new StyleParagraphDto() ? null : dto;
    }

    /// <summary>A relative size becomes the points it amounts to over the parent.</summary>
    private static StyleCharacterDto? CharacterOf(OdtStyles styles, OdtStyle style)
    {
        var own = CharacterFormat.Of(style.Properties("text-properties"), styles.Fonts);
        if (own.FontSizePt is null && own.FontSizeFactor is not null)
        {
            var chain = styles.Chain(style.Family, style.Name);
            var under = styles.Character(chain.Take(chain.Count - 1), styles.DefaultCharacter());
            own = under.Over(own) is { FontSizePt: { } size } ? own with { FontSizePt = size } : own;
        }

        var dto = own.ToDto();
        return dto == new StyleCharacterDto() ? null : dto;
    }

    /// <summary>ODF counts outline levels from 1 to 10; Word from 0 to 8.</summary>
    private static int? WordOutlineLevelOf(int? level) =>
        level is >= 1 and <= WordOutlineLevels ? level - 1 : null;

    /// <summary>LibreOffice escapes the name: <c>Text_20_body</c> is "Text body".</summary>
    public static string DisplayNameOf(string name) =>
        EscapedCharacter().Replace(name, match =>
        {
            var code = int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return code is > 0 and <= MaxCodePoint and not (>= SurrogateFirst and <= SurrogateLast)
                ? char.ConvertFromUtf32(code)
                : match.Value;
        });

    private const int MaxCodePoint = 0x10FFFF;
    private const int SurrogateFirst = 0xD800;
    private const int SurrogateLast = 0xDFFF;

    [GeneratedRegex("_([0-9a-fA-F]{2,6})_")]
    private static partial Regex EscapedCharacter();
}

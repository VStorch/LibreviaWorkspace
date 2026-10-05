using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Line spacing as the file declares it: a multiple or points. The TS side (<c>line-metrics.ts</c>)
/// multiplies by the font's natural height, after the cascade.
/// </summary>
public sealed record LineSpacingDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("factor")] double? Factor = null,
    [property: JsonPropertyName("pt")] double? Points = null);

/// <c>null</c> means "the style does not say", and lets the inherited value through; zero means "no
/// space".
public sealed record StyleParagraphDto(
    [property: JsonPropertyName("textAlign")] string? TextAlign = null,
    [property: JsonPropertyName("indentMm")] double? IndentMm = null,
    [property: JsonPropertyName("indentRightMm")] double? IndentRightMm = null,
    [property: JsonPropertyName("firstLineMm")] double? FirstLineMm = null,
    [property: JsonPropertyName("spaceBefore")] double? SpaceBefore = null,
    [property: JsonPropertyName("spaceAfter")] double? SpaceAfter = null,
    [property: JsonPropertyName("lineSpacing")] LineSpacingDto? LineSpacing = null,
    [property: JsonPropertyName("keepNext")] bool? KeepNext = null,
    [property: JsonPropertyName("keepLines")] bool? KeepLines = null,
    [property: JsonPropertyName("pageBreakBefore")] bool? PageBreakBefore = null,
    [property: JsonPropertyName("contextualSpacing")] bool? ContextualSpacing = null,
    [property: JsonPropertyName("outlineLevel")] int? OutlineLevel = null,
    [property: JsonPropertyName("background")] string? Background = null,
    [property: JsonPropertyName("widowControl")] bool? WidowControl = null);

public sealed record StyleCharacterDto(
    [property: JsonPropertyName("fontFamily")] string? FontFamily = null,
    [property: JsonPropertyName("fontSize")] string? FontSize = null,
    [property: JsonPropertyName("bold")] bool? Bold = null,
    [property: JsonPropertyName("italic")] bool? Italic = null,
    [property: JsonPropertyName("underline")] bool? Underline = null,
    [property: JsonPropertyName("strike")] bool? Strike = null,
    [property: JsonPropertyName("allCaps")] bool? AllCaps = null,
    [property: JsonPropertyName("smallCaps")] bool? SmallCaps = null,
    [property: JsonPropertyName("verticalAlign")] string? VerticalAlign = null,
    [property: JsonPropertyName("color")] string? Color = null,
    [property: JsonPropertyName("highlight")] string? Highlight = null);

/// <summary>As the file declares it.</summary>
public sealed record StyleDefinitionDto(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("qFormat")] bool QFormat,
    [property: JsonPropertyName("hidden")] bool Hidden,
    [property: JsonPropertyName("custom")] bool Custom,
    [property: JsonPropertyName("basedOn")] string? BasedOn = null,
    [property: JsonPropertyName("next")] string? Next = null,
    [property: JsonPropertyName("link")] string? Link = null,
    [property: JsonPropertyName("uiPriority")] int? UiPriority = null,
    [property: JsonPropertyName("paragraph")] StyleParagraphDto? Paragraph = null,
    [property: JsonPropertyName("character")] StyleCharacterDto? Character = null);

/// <c>w:docDefaults</c>, and the style that applies when a paragraph names none.
public sealed record StyleDefaultsDto(
    [property: JsonPropertyName("paragraph")] StyleParagraphDto Paragraph,
    [property: JsonPropertyName("character")] StyleCharacterDto Character,
    [property: JsonPropertyName("paragraphStyleId")] string? ParagraphStyleId,
    [property: JsonPropertyName("characterStyleId")] string? CharacterStyleId);

public sealed record StyleSheetDto(
    [property: JsonPropertyName("defaults")] StyleDefaultsDto Defaults,
    [property: JsonPropertyName("styles")] Dictionary<string, StyleDefinitionDto> Styles);

/// <c>word/styles.xml</c> as **data**, without inheritance, in editor units (line spacing is the
/// exception, see <see cref="LineSpacingDto"/>); the cascade is on the TS side. Left out: table and
/// numbering styles, latent ones, tabs, borders, frame, language and <c>w:rsid</c>:
/// <c>styles.xml</c> goes back untouched, and they just do not show in the panel.
public static class StyleReader
{
    private const int MaxStyles = 4000;

    public static StyleSheetDto Read(MainDocumentPart part)
    {
        var fonts = new FontTable(part);
        var styles = part.StyleDefinitionsPart?.Styles;

        var definitions = new Dictionary<string, StyleDefinitionDto>(StringComparer.Ordinal);
        string? defaultParagraph = null;
        string? defaultCharacter = null;

        foreach (var style in styles?.Elements<Style>() ?? [])
        {
            if (style.StyleId?.Value is not { Length: > 0 } id) continue;
            if (KindOf(style.Type?.Value) is not { } kind) continue;
            // A repeated id: the first wins, as in the resolver.
            if (definitions.ContainsKey(id) || definitions.Count >= MaxStyles) continue;

            definitions[id] = Definition(style, id, kind, fonts);

            if (style.Default?.Value != true) continue;
            if (kind == "paragraph") defaultParagraph ??= id;
            else defaultCharacter ??= id;
        }

        var defaults = new StyleDefaultsDto(
            ParagraphOf(styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle)
                ?? new StyleParagraphDto(),
            CharacterOf(styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle, fonts)
                ?? new StyleCharacterDto(),
            defaultParagraph,
            defaultCharacter);

        return new StyleSheetDto(defaults, definitions);
    }

    /// <summary>
    /// Only paragraph and character; a missing type is paragraph, as the schema says.
    /// </summary>
    private static string? KindOf(StyleValues? type)
    {
        if (type is null) return "paragraph";
        if (type.Value == StyleValues.Paragraph) return "paragraph";
        if (type.Value == StyleValues.Character) return "character";
        return null;
    }

    private static StyleDefinitionDto Definition(Style style, string id, string kind, FontTable fonts)
    {
        // Without `w:name`, the id.
        var name = style.StyleName?.Val?.Value is { Length: > 0 } declared ? declared : id;

        return new StyleDefinitionDto(
            id,
            name,
            kind,
            QFormat: IsOn(style.PrimaryStyle),
            // `w:hidden` always hides and `w:semiHidden` until used: the panel only asks whether it
            // shows.
            Hidden: IsOn(style.StyleHidden) || IsOn(style.SemiHidden),
            Custom: style.CustomStyle?.Value == true,
            BasedOn: Text(style.BasedOn?.Val?.Value),
            Next: Text(style.NextParagraphStyle?.Val?.Value),
            Link: Text(style.LinkedStyle?.Val?.Value),
            UiPriority: style.UIPriority?.Val?.Value,
            Paragraph: ParagraphOf(style.StyleParagraphProperties),
            Character: CharacterOf(style.StyleRunProperties, fonts));
    }

    /// <summary>
    /// Through <c>GetFirstChild</c>: style, defaults and paragraph have three classes for the same
    /// list.
    /// </summary>
    private static StyleParagraphDto? ParagraphOf(OpenXmlElement? properties)
    {
        if (properties is null) return null;

        var spacing = properties.GetFirstChild<SpacingBetweenLines>();
        var indentation = properties.GetFirstChild<Indentation>();

        var dto = new StyleParagraphDto(
            TextAlign: AlignmentOf(properties.GetFirstChild<Justification>()?.Val),
            IndentMm: Millimeters(indentation?.Left?.Value),
            IndentRightMm: Millimeters(indentation?.Right?.Value),
            FirstLineMm: FirstLineOf(indentation),
            SpaceBefore: Points(spacing?.Before?.Value),
            SpaceAfter: Points(spacing?.After?.Value),
            LineSpacing: LineSpacingOf(spacing),
            KeepNext: Toggle(properties.GetFirstChild<KeepNext>()),
            KeepLines: Toggle(properties.GetFirstChild<KeepLines>()),
            PageBreakBefore: Toggle(properties.GetFirstChild<PageBreakBefore>()),
            ContextualSpacing: Toggle(properties.GetFirstChild<ContextualSpacing>()),
            OutlineLevel: properties.GetFirstChild<OutlineLevel>()?.Val?.Value,
            Background: ShadingOf(properties.GetFirstChild<Shading>()),
            WidowControl: Toggle(properties.GetFirstChild<WidowControl>()));

        // "Declares nothing I can read" and "does not exist" amount to the same for whoever
        // inherits.
        return dto == new StyleParagraphDto() ? null : dto;
    }

    private static StyleCharacterDto? CharacterOf(OpenXmlElement? properties, FontTable fonts)
    {
        if (properties is null) return null;

        var runFonts = properties.GetFirstChild<RunFonts>();
        var font = runFonts?.Ascii?.Value ?? runFonts?.HighAnsi?.Value;
        var underline = properties.GetFirstChild<Underline>()?.Val;

        var dto = new StyleCharacterDto(
            // With the generic substitute behind it, as in the body reader.
            FontFamily: string.IsNullOrWhiteSpace(font) ? null : fonts.Stack(font),
            FontSize: PointsCss(properties.GetFirstChild<FontSize>()?.Val?.Value),
            Bold: Toggle(properties.GetFirstChild<Bold>()),
            Italic: Toggle(properties.GetFirstChild<Italic>()),
            // `w:u` carries the line style, and `none` turns it off.
            Underline: underline is null ? null : underline.Value != UnderlineValues.None,
            Strike: Toggle(properties.GetFirstChild<Strike>()),
            AllCaps: Toggle(properties.GetFirstChild<Caps>()),
            SmallCaps: Toggle(properties.GetFirstChild<SmallCaps>()),
            VerticalAlign: VerticalOf(properties.GetFirstChild<VerticalTextAlignment>()?.Val),
            Color: RunReader.ColorOf(properties.GetFirstChild<Color>()?.Val),
            Highlight: RunReader.HighlightOf(properties));

        return dto == new StyleCharacterDto() ? null : dto;
    }

    /// <summary>
    /// Absent is silence: turning off the parent's bold is not the same as not mentioning it.
    /// </summary>
    private static bool? Toggle(OnOffType? toggle) => toggle is null ? null : RunReader.IsOn(toggle);

    /// <c>w:qFormat</c>, <c>w:hidden</c> and <c>w:semiHidden</c> have their own type: reading
    /// presence would make <c>w:semiHidden w:val="off"</c> hide the style.
    private static bool IsOn(OnOffOnlyType? toggle) =>
        toggle is not null && (toggle.Val is null || toggle.Val.Value == OnOffOnlyValues.On);

    private static string? Text(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    private static string? AlignmentOf(EnumValue<JustificationValues>? value)
    {
        if (value is null) return null;
        if (value == JustificationValues.Center) return "center";
        if (value == JustificationValues.Right) return "right";
        if (value == JustificationValues.Both || value == JustificationValues.Distribute) return "justify";
        return "left";
    }

    private static string? VerticalOf(EnumValue<VerticalPositionValues>? value)
    {
        if (value is null) return null;
        if (value.Value == VerticalPositionValues.Superscript) return "super";
        if (value.Value == VerticalPositionValues.Subscript) return "sub";
        return null;
    }

    /// <c>w:firstLine</c> pushes, <c>w:hanging</c> pulls.
    private static double? FirstLineOf(Indentation? indentation)
    {
        if (Millimeters(indentation?.FirstLine?.Value) is { } firstLine and > 0) return firstLine;
        if (Millimeters(indentation?.Hanging?.Value) is { } hanging and > 0) return -hanging;
        return null;
    }

    private static double? Millimeters(string? twips) =>
        int.TryParse(twips, out var value) ? Math.Round(Unit.TwipsToMillimeters(value), 2) : null;

    /// <summary>
    /// Twips → points with two exact decimals (1 twip = 0.05 pt): the measure round-trips. An
    /// explicit zero counts.
    /// </summary>
    private static double? Points(string? twips) =>
        int.TryParse(twips, out var value) && value >= 0 ? Math.Round(Unit.TwipsToPoints(value), 2) : null;

    private static string? PointsCss(string? halfPoints) =>
        double.TryParse(halfPoints, out var value) && value > 0
            ? RunReader.FormatPoints(Unit.HalfPointsToPoints(value))
            : null;

    /// <summary>
    /// Unmultiplied; the multiple to four decimals, the grid it goes back to the file on.
    /// </summary>
    private static LineSpacingDto? LineSpacingOf(SpacingBetweenLines? spacing)
    {
        if (!int.TryParse(spacing?.Line?.Value, out var value) || value <= 0) return null;

        var rule = spacing!.LineRule?.Value;
        if (rule is not null && rule != LineSpacingRuleValues.Auto)
        {
            var points = Math.Round(Unit.TwipsToPoints(value), 2);
            return rule == LineSpacingRuleValues.Exact
                ? new LineSpacingDto("exact", Points: points)
                : new LineSpacingDto("atLeast", Points: points);
        }

        return new LineSpacingDto("multiple", Factor: Math.Round(value / 240.0, 4));
    }

    /// <summary>"auto" is not a color.</summary>
    private static string? ShadingOf(Shading? shading) => RunReader.ColorOf(shading?.Fill);
}

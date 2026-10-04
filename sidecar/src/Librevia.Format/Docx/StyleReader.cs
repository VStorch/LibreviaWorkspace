using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Entrelinha como o arquivo a declara: múltiplo ou pontos. Quem multiplica pela
/// altura natural da fonte é o lado TS (<c>line-metrics.ts</c>), depois da cascata.
/// </summary>
public sealed record LineSpacingDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("factor")] double? Factor = null,
    [property: JsonPropertyName("pt")] double? Points = null);

/// <summary><c>null</c> é "o estilo não fala disso", e deixa o herdado passar; zero é "nenhum espaço".</summary>
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

/// <summary>Como o arquivo o declara.</summary>
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

/// <summary>O <c>w:docDefaults</c>, e quem vale quando o parágrafo não diz estilo.</summary>
public sealed record StyleDefaultsDto(
    [property: JsonPropertyName("paragraph")] StyleParagraphDto Paragraph,
    [property: JsonPropertyName("character")] StyleCharacterDto Character,
    [property: JsonPropertyName("paragraphStyleId")] string? ParagraphStyleId,
    [property: JsonPropertyName("characterStyleId")] string? CharacterStyleId);

public sealed record StyleSheetDto(
    [property: JsonPropertyName("defaults")] StyleDefaultsDto Defaults,
    [property: JsonPropertyName("styles")] Dictionary<string, StyleDefinitionDto> Styles);

/// <summary>
/// <c>word/styles.xml</c> como **dado**, sem herança, nas unidades do editor (a
/// entrelinha é a exceção, ver <see cref="LineSpacingDto"/>); a cascata é do lado
/// TS. Ficam de fora estilos de tabela e de numeração, os latentes, tabulações,
/// bordas, moldura, idioma e <c>w:rsid</c>: <c>styles.xml</c> volta intocado, e eles
/// só não aparecem no painel.
/// </summary>
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
            // Id repetido: vale o primeiro, como no resolvedor.
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

    /// <summary>Só parágrafo e caractere; tipo ausente é parágrafo, como diz o esquema.</summary>
    private static string? KindOf(StyleValues? type)
    {
        if (type is null) return "paragraph";
        if (type.Value == StyleValues.Paragraph) return "paragraph";
        if (type.Value == StyleValues.Character) return "character";
        return null;
    }

    private static StyleDefinitionDto Definition(Style style, string id, string kind, FontTable fonts)
    {
        // Sem `w:name`, o id.
        var name = style.StyleName?.Val?.Value is { Length: > 0 } declared ? declared : id;

        return new StyleDefinitionDto(
            id,
            name,
            kind,
            QFormat: IsOn(style.PrimaryStyle),
            // `w:hidden` esconde sempre e `w:semiHidden` até ser usado: o painel pergunta só se aparece.
            Hidden: IsOn(style.StyleHidden) || IsOn(style.SemiHidden),
            Custom: style.CustomStyle?.Value == true,
            BasedOn: Text(style.BasedOn?.Val?.Value),
            Next: Text(style.NextParagraphStyle?.Val?.Value),
            Link: Text(style.LinkedStyle?.Val?.Value),
            UiPriority: style.UIPriority?.Val?.Value,
            Paragraph: ParagraphOf(style.StyleParagraphProperties),
            Character: CharacterOf(style.StyleRunProperties, fonts));
    }

    /// <summary>Por <c>GetFirstChild</c>: estilo, padrão e parágrafo têm três classes para a mesma lista.</summary>
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

        // "Não declara nada que eu saiba ler" e "não existe" dão no mesmo para quem herda.
        return dto == new StyleParagraphDto() ? null : dto;
    }

    private static StyleCharacterDto? CharacterOf(OpenXmlElement? properties, FontTable fonts)
    {
        if (properties is null) return null;

        var runFonts = properties.GetFirstChild<RunFonts>();
        var font = runFonts?.Ascii?.Value ?? runFonts?.HighAnsi?.Value;
        var underline = properties.GetFirstChild<Underline>()?.Val;

        var dto = new StyleCharacterDto(
            // Com a substituta genérica atrás, como no leitor do corpo.
            FontFamily: string.IsNullOrWhiteSpace(font) ? null : fonts.Stack(font),
            FontSize: PointsCss(properties.GetFirstChild<FontSize>()?.Val?.Value),
            Bold: Toggle(properties.GetFirstChild<Bold>()),
            Italic: Toggle(properties.GetFirstChild<Italic>()),
            // `w:u` carrega o estilo do risco, e `none` desliga.
            Underline: underline is null ? null : underline.Value != UnderlineValues.None,
            Strike: Toggle(properties.GetFirstChild<Strike>()),
            AllCaps: Toggle(properties.GetFirstChild<Caps>()),
            SmallCaps: Toggle(properties.GetFirstChild<SmallCaps>()),
            VerticalAlign: VerticalOf(properties.GetFirstChild<VerticalTextAlignment>()?.Val),
            Color: RunReader.ColorOf(properties.GetFirstChild<Color>()?.Val),
            Highlight: RunReader.HighlightOf(properties));

        return dto == new StyleCharacterDto() ? null : dto;
    }

    /// <summary>Ausente é silêncio: desligar o negrito do pai não é o mesmo que não falar dele.</summary>
    private static bool? Toggle(OnOffType? toggle) => toggle is null ? null : RunReader.IsOn(toggle);

    /// <summary>
    /// <c>w:qFormat</c>, <c>w:hidden</c> e <c>w:semiHidden</c> têm tipo próprio: ler a presença
    /// faria <c>w:semiHidden w:val="off"</c> esconder o estilo.
    /// </summary>
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

    /// <summary><c>w:firstLine</c> empurra, <c>w:hanging</c> puxa.</summary>
    private static double? FirstLineOf(Indentation? indentation)
    {
        if (Millimeters(indentation?.FirstLine?.Value) is { } firstLine and > 0) return firstLine;
        if (Millimeters(indentation?.Hanging?.Value) is { } hanging and > 0) return -hanging;
        return null;
    }

    private static double? Millimeters(string? twips) =>
        int.TryParse(twips, out var value) ? Math.Round(value * 25.4 / 1440, 2) : null;

    /// <summary>Twips → pontos com duas casas, exatas (1 twip = 0,05 pt): a medida vai e volta. Zero explícito conta.</summary>
    private static double? Points(string? twips) =>
        int.TryParse(twips, out var value) && value >= 0 ? Math.Round(value / 20.0, 2) : null;

    /// <summary><c>w:sz</c> vem em meios-pontos.</summary>
    private static string? PointsCss(string? halfPoints) =>
        double.TryParse(halfPoints, out var value) && value > 0
            ? RunReader.FormatPoints(value / 2)
            : null;

    /// <summary>Sem multiplicar; o múltiplo a quatro casas, a grade em que volta ao arquivo.</summary>
    private static LineSpacingDto? LineSpacingOf(SpacingBetweenLines? spacing)
    {
        if (!int.TryParse(spacing?.Line?.Value, out var value) || value <= 0) return null;

        var rule = spacing!.LineRule?.Value;
        if (rule is not null && rule != LineSpacingRuleValues.Auto)
        {
            var points = Math.Round(value / 20.0, 2);
            return rule == LineSpacingRuleValues.Exact
                ? new LineSpacingDto("exact", Points: points)
                : new LineSpacingDto("atLeast", Points: points);
        }

        return new LineSpacingDto("multiple", Factor: Math.Round(value / 240.0, 4));
    }

    /// <summary>"auto" não é cor.</summary>
    private static string? ShadingOf(Shading? shading) => RunReader.ColorOf(shading?.Fill);
}

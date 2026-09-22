using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>Entrelinha na forma bruta do arquivo: múltiplo, ou medida em pontos.</summary>
/// <remarks>
/// **Não** é o número do CSS que o bloco carrega. O múltiplo do OOXML é medido
/// sobre a altura natural da fonte, e com herança a fonte pode vir do estilo —
/// então quem multiplica é o lado TS, onde `line-metrics.ts` já sabe a altura de
/// cada família. Emitir o número já multiplicado obrigaria a resolver a cascata
/// aqui, e ela é da entrega seguinte.
/// </remarks>
public sealed record LineSpacingDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("factor")] double? Factor = null,
    [property: JsonPropertyName("pt")] double? Points = null);

/// <summary>Propriedades de parágrafo de um estilo, nas unidades do editor.</summary>
/// <remarks>
/// Tudo anulável de propósito: <c>null</c> é "o estilo não fala disso", e é o
/// silêncio que a cascata precisa para deixar o valor herdado passar. Zero é
/// outra coisa — é o estilo dizendo "nenhum espaço aqui".
/// </remarks>
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
    [property: JsonPropertyName("background")] string? Background = null);

/// <summary>Propriedades de caractere de um estilo, nas unidades do editor.</summary>
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

/// <summary>Um estilo do documento, como o arquivo o declara.</summary>
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

/// <summary>O `w:docDefaults`, e quem vale quando o parágrafo não diz estilo.</summary>
public sealed record StyleDefaultsDto(
    [property: JsonPropertyName("paragraph")] StyleParagraphDto Paragraph,
    [property: JsonPropertyName("character")] StyleCharacterDto Character,
    [property: JsonPropertyName("paragraphStyleId")] string? ParagraphStyleId,
    [property: JsonPropertyName("characterStyleId")] string? CharacterStyleId);

/// <summary>Os estilos do documento, por id.</summary>
public sealed record StyleSheetDto(
    [property: JsonPropertyName("defaults")] StyleDefaultsDto Defaults,
    [property: JsonPropertyName("styles")] Dictionary<string, StyleDefinitionDto> Styles);

/// <summary>
/// Lê `word/styles.xml` como **dado**, nas unidades do editor.
/// </summary>
/// <remarks>
/// É o outro caminho, ao lado do <see cref="StyleResolver"/>: ele resolve a
/// cascata e achata o resultado no bloco, para a tela desenhar hoje; este traz as
/// definições como o arquivo as declara, sem herança nenhuma — cada estilo diz só
/// o que ele mesmo diz, e o `basedOn` fica registrado para quem for resolver.
/// Quem resolve é o lado TS, na entrega seguinte.
///
/// <b>Unidades do editor</b>: pontos no espaçamento e no tamanho da fonte,
/// milímetros no recuo, pilha de CSS na família, hexadecimal na cor. São as
/// mesmas dos atributos do bloco, para que a tela não precise de duas traduções.
/// A entrelinha é a exceção declarada: sai em forma bruta (ver
/// <see cref="LineSpacingDto"/>).
///
/// <b>O que fica de fora</b>, e por que isso não é perda: estilos de tabela e de
/// numeração (o modelo não representa borda nem margem de célula), a lista de
/// estilos latentes (`w:latentStyles`, que é galeria do Word e não formatação),
/// numeração (`w:numPr`), paradas de tabulação, bordas de parágrafo, moldura
/// (`w:framePr`), idioma, espaçamento entre letras e os carimbos de revisão
/// (`w:rsid`). Nada disso se perde do arquivo: `word/styles.xml` é parte intocada
/// na gravação cirúrgica e volta byte a byte — o que não está aqui simplesmente
/// não aparece no painel de estilos. O dia em que a tela passar a desenhar a
/// partir destes dados, o que faltar precisa entrar antes, e não depois.
/// </remarks>
public static class StyleReader
{
    /// <summary>Teto de segurança: nenhum documento de verdade passa daqui.</summary>
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
            // Id repetido é arquivo malformado. Vale o primeiro, como no
            // resolvedor: dois estilos com o mesmo id não podem virar um só.
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
    /// Só parágrafo e caractere: os outros dois tipos não têm onde morar no modelo.
    /// </summary>
    /// <remarks>
    /// Tipo ausente é parágrafo — é o que o esquema do OOXML diz, e o que o
    /// resolvedor já assume ao procurar o estilo padrão.
    /// </remarks>
    private static string? KindOf(StyleValues? type)
    {
        if (type is null) return "paragraph";
        if (type.Value == StyleValues.Paragraph) return "paragraph";
        if (type.Value == StyleValues.Character) return "character";
        return null;
    }

    private static StyleDefinitionDto Definition(Style style, string id, string kind, FontTable fonts)
    {
        // Sem `w:name` o id é o melhor nome disponível: um estilo sem nome é
        // arquivo malformado, e mostrar a lista sem ele seria pior.
        var name = style.StyleName?.Val?.Value is { Length: > 0 } declared ? declared : id;

        return new StyleDefinitionDto(
            id,
            name,
            kind,
            QFormat: IsOn(style.PrimaryStyle),
            // Os dois jeitos de o Word esconder um estilo da galeria juntos: o
            // `w:hidden` esconde sempre e o `w:semiHidden` esconde até ser usado.
            // A pergunta que o painel faz é uma só — "isto aparece na lista?".
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
    /// As propriedades de parágrafo de um `w:pPr`, seja de estilo ou de padrão.
    /// </summary>
    /// <remarks>
    /// Por <c>GetFirstChild</c>, e não pelas propriedades tipadas: o OOXML tem
    /// três classes diferentes para a mesma lista de filhos — a do estilo, a do
    /// padrão do documento e a do parágrafo — e ler por elemento serve às três
    /// sem copiar este método três vezes.
    /// </remarks>
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
            Background: ShadingOf(properties.GetFirstChild<Shading>()));

        // Um `w:pPr` que só tem coisas que não lemos não vira objeto vazio no
        // modelo: "não declara nada que eu saiba ler" e "não existe" dão no mesmo
        // para quem herda.
        return dto == new StyleParagraphDto() ? null : dto;
    }

    private static StyleCharacterDto? CharacterOf(OpenXmlElement? properties, FontTable fonts)
    {
        if (properties is null) return null;

        var runFonts = properties.GetFirstChild<RunFonts>();
        var font = runFonts?.Ascii?.Value ?? runFonts?.HighAnsi?.Value;
        var underline = properties.GetFirstChild<Underline>()?.Val;

        var dto = new StyleCharacterDto(
            // Com a substituta genérica atrás, como no leitor do corpo: a fonte
            // que o documento pede pode não existir na máquina de quem abre.
            FontFamily: string.IsNullOrWhiteSpace(font) ? null : fonts.Stack(font),
            FontSize: PointsCss(properties.GetFirstChild<FontSize>()?.Val?.Value),
            Bold: Toggle(properties.GetFirstChild<Bold>()),
            Italic: Toggle(properties.GetFirstChild<Italic>()),
            // `w:u` não é alternância: carrega o estilo do risco, e `none` desliga.
            Underline: underline is null ? null : underline.Value != UnderlineValues.None,
            Strike: Toggle(properties.GetFirstChild<Strike>()),
            AllCaps: Toggle(properties.GetFirstChild<Caps>()),
            SmallCaps: Toggle(properties.GetFirstChild<SmallCaps>()),
            VerticalAlign: VerticalOf(properties.GetFirstChild<VerticalTextAlignment>()?.Val),
            Color: RunReader.ColorOf(properties.GetFirstChild<Color>()?.Val),
            Highlight: RunReader.HighlightOf(properties));

        return dto == new StyleCharacterDto() ? null : dto;
    }

    /// <summary>Elemento ausente é silêncio; presente é o valor dele.</summary>
    /// <remarks>
    /// A diferença importa: um estilo que **desliga** o negrito do estilo pai não
    /// é o mesmo que um estilo que não fala de negrito.
    /// </remarks>
    private static bool? Toggle(OnOffType? toggle) => toggle is null ? null : RunReader.IsOn(toggle);

    /// <summary>
    /// A alternância que o `w:style` usa, que é outra classe da mesma ideia.
    /// </summary>
    /// <remarks>
    /// `w:qFormat`, `w:hidden` e `w:semiHidden` só aceitam `on` e `off`, e o SDK
    /// lhes dá um tipo próprio. Ler a presença do elemento — o erro clássico —
    /// faria um `w:semiHidden w:val="off"` esconder o estilo que ele revela.
    /// </remarks>
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

    /// <summary>A primeira linha anda para os dois lados: `w:firstLine` empurra, `w:hanging` puxa.</summary>
    private static double? FirstLineOf(Indentation? indentation)
    {
        if (Millimeters(indentation?.FirstLine?.Value) is { } firstLine and > 0) return firstLine;
        if (Millimeters(indentation?.Hanging?.Value) is { } hanging and > 0) return -hanging;
        return null;
    }

    /// <summary>Twips → milímetros, na precisão que a interface mostra.</summary>
    private static double? Millimeters(string? twips) =>
        int.TryParse(twips, out var value) ? Math.Round(value * 25.4 / 1440, 2) : null;

    /// <summary>Twips → pontos. Zero explícito é preservado: é uma instrução.</summary>
    /// <remarks>
    /// Duas casas, e não uma como no leitor do corpo: 1 twip vale 0,05 pt, então
    /// duas casas são **exatas** e uma perde um quarto de ponto. Aqui isso
    /// importa porque a medida vai e volta — o espaço de 14,75 pt do título do
    /// documento novo virava 14,8, e o teste de contrato com a tabela de dados
    /// acusava uma diferença que o arquivo não tem.
    /// </remarks>
    private static double? Points(string? twips) =>
        int.TryParse(twips, out var value) && value >= 0 ? Math.Round(value / 20.0, 2) : null;

    /// <summary>`w:sz` vem em meios-pontos: 24 significa 12 pt.</summary>
    private static string? PointsCss(string? halfPoints) =>
        double.TryParse(halfPoints, out var value) && value > 0
            ? RunReader.FormatPoints(value / 2)
            : null;

    /// <summary>
    /// A entrelinha como o arquivo a declara, sem multiplicar nada.
    /// </summary>
    /// <remarks>
    /// O múltiplo vem em 240-avos, e é arredondado a quatro casas — a mesma grade
    /// em que ele volta ao arquivo. Guardar mais casas faria o valor lido divergir
    /// do declarado sem que ninguém tivesse mudado nada.
    /// </remarks>
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

    /// <summary>Fundo do parágrafo, quando é cor de verdade — "auto" não é.</summary>
    private static string? ShadingOf(Shading? shading) => RunReader.ColorOf(shading?.Fill);
}

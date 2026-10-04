using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

public sealed record PageSetupDto(
    [property: JsonPropertyName("size")] string Size,
    [property: JsonPropertyName("orientation")] string Orientation,
    [property: JsonPropertyName("margins")] MarginsDto Margins,
    [property: JsonPropertyName("headerBand")] BandDto? Header,
    [property: JsonPropertyName("footerBand")] BandDto? Footer,
    // Campos novos e opcionais, para o `.sdoc` gravado sem eles continuar abrindo.
    [property: JsonPropertyName("firstHeaderBand")] BandDto? FirstHeader = null,
    [property: JsonPropertyName("firstFooterBand")] BandDto? FirstFooter = null,
    [property: JsonPropertyName("evenHeaderBand")] BandDto? EvenHeader = null,
    [property: JsonPropertyName("evenFooterBand")] BandDto? EvenFooter = null,
    /// <summary><c>w:pgMar/@header</c> e <c>@footer</c>: a origem vertical das âncoras de dentro da faixa.</summary>
    [property: JsonPropertyName("headerDistanceMm")] double HeaderDistanceMm = 12.5,
    [property: JsonPropertyName("footerDistanceMm")] double FooterDistanceMm = 12.5,
    // A faixa de texto simples do documento novo, com `{n}` e `{total}` (PlainBandWriter).
    [property: JsonPropertyName("header")] string? HeaderText = null,
    [property: JsonPropertyName("footer")] string? FooterText = null,
    // `w:pgNumType` e os interruptores das faixas: ausente é "não mexa", e não "desligue".
    [property: JsonPropertyName("pageNumberFormat")] string? PageNumberFormat = null,
    // Ausente é "não mexa no `w:start`"; nulo é "sem início". `JsonElement` distingue os dois.
    [property: JsonPropertyName("pageNumberStart")] System.Text.Json.JsonElement PageNumberStart = default,
    [property: JsonPropertyName("titlePage")] bool? TitlePage = null,
    [property: JsonPropertyName("evenAndOddHeaders")] bool? EvenAndOddHeaders = null,
    // `Id` só nas seções antes da última: o `sectionBreak` do parágrafo que as encerra.
    [property: JsonPropertyName("id")] string? Id = null,
    // `w:sectPr/w:type`. Ausente é "não mexa".
    [property: JsonPropertyName("start")] string? Start = null,
    // `w:cols`. Ausente é "não mexa".
    [property: JsonPropertyName("columns")] ColumnsDto? Columns = null);

/// <summary>
/// <c>w:cols</c>. Larguras diferentes vêm em <c>WidthsMm</c>: a tela desenha iguais, e o
/// arquivo as mantém enquanto o número de colunas não mudar.
/// </summary>
public sealed record ColumnsDto(
    [property: JsonPropertyName("count")] int Count,
    [property: JsonPropertyName("spaceMm")] double SpaceMm,
    [property: JsonPropertyName("separator")] bool Separator,
    [property: JsonPropertyName("widthsMm")] List<double>? WidthsMm = null);

public sealed record MarginsDto(
    [property: JsonPropertyName("top")] double Top,
    [property: JsonPropertyName("right")] double Right,
    [property: JsonPropertyName("bottom")] double Bottom,
    [property: JsonPropertyName("left")] double Left);

/// <summary><c>w:sectPr</c> → configuração de página.</summary>
public static class PageReader
{
    private const double TwipsPerMillimeter = 1440 / 25.4;

    private const double EmusPerTwip = 914400.0 / 1440;

    /// <summary>Meio milímetro: o mesmo A4 do LibreOffice e do Word difere no último twip.</summary>
    private const int PaperTolerance = 30;

    /// <summary>Em twips, numa tabela só, para um A4 não ter duas larguras.</summary>
    private static readonly (string Name, uint Short, uint Long)[] Papers =
    [
        ("A4", 11906U, 16838U),
        ("Letter", 12240U, 15840U),
    ];

    /// <summary>A4 para um nome que a tabela não tem.</summary>
    public static (uint Short, uint Long) TwipsOfPaper(string name)
    {
        foreach (var paper in Papers)
        {
            if (string.Equals(paper.Name, name, StringComparison.Ordinal)) return (paper.Short, paper.Long);
        }

        return (Papers[0].Short, Papers[0].Long);
    }

    public static (double Short, double Long) MillimetersOfPaper(string name)
    {
        var (shortSide, longSide) = TwipsOfPaper(name);
        return (shortSide / TwipsPerMillimeter, longSide / TwipsPerMillimeter);
    }

    /// <summary>
    /// A última seção é a "página" do modelo; as anteriores vêm em ordem, com o id
    /// que o leitor põe no parágrafo que as encerra (<see cref="SectionIds"/>). Cada
    /// uma leva só as faixas que **declara**: a herança é de quem desenha.
    /// </summary>
    public static (PageSetupDto Page, List<PageSetupDto>? Sections) ReadAll(
        Body body,
        MainDocumentPart part,
        Inventory inventory)
    {
        var all = body.Descendants<SectionProperties>().ToList();
        if (all.Count == 0) return (Default(), null);

        // O documento sem `w:sectPr` no corpo é inválido, mas acontece: a última de parágrafo faz as vezes.
        var ids = SectionIds(body);
        var earlier = new List<PageSetupDto>();
        for (var index = 0; index < all.Count - 1; index++)
        {
            var section = all[index];
            earlier.Add(ReadOne(section, part, inventory, index) with
            {
                Id = ids.TryGetValue(section, out var id) ? id : $"s{index + 1}",
            });
        }

        var page = ReadOne(all[^1], part, inventory, all.Count - 1);
        return (page, earlier.Count == 0 ? null : earlier);
    }

    /// <summary><c>s1</c>, <c>s2</c>…, posicional como o <c>oid</c>: a leitura de referência reproduz os ids.</summary>
    public static Dictionary<SectionProperties, string> SectionIds(Body body)
    {
        var ids = new Dictionary<SectionProperties, string>(ReferenceEqualityComparer.Instance);
        var all = body.Descendants<SectionProperties>().ToList();
        for (var index = 0; index < all.Count - 1; index++) ids[all[index]] = $"s{index + 1}";
        return ids;
    }

    private static PageSetupDto ReadOne(SectionProperties section, MainDocumentPart part, Inventory inventory, int index)
    {
        // A tela mostra decimal e o arquivo continua pedindo o dele: avisa-se.
        if (section.GetFirstChild<PageNumberType>()?.Format?.InnerText is { } pageFormat &&
            !PageNumberFormats.Contains(pageFormat))
        {
            inventory.NoteInvisible(
                $"formato de número de página \"{pageFormat}\" (mostrado em algarismos; o arquivo o mantém)");
        }
        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        var margin = section.GetFirstChild<PageMargin>();

        var landscape = size?.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        // Sem `w:pgSz`, as medidas do primeiro papel da tabela.
        var widthTwips = (double?)size?.Width?.Value ?? Papers[0].Short;
        var heightTwips = (double?)size?.Height?.Value ?? Papers[0].Long;

        // Papel fora de A4 e Carta aparece como o mais próximo, e o arquivo mantém a medida: invisibilidade.
        if (!IsKnownPaper(widthTwips, heightTwips))
        {
            inventory.NoteInvisible(
                "o tamanho do papel deste documento não é A4 nem Carta (ele é preservado no arquivo)");
        }

        // É contra a largura da coluna que a faixa decide o terço de cada peça.
        var contentWidthEmus = Math.Max(
            (widthTwips - (margin?.Left?.Value ?? 1440) - (margin?.Right?.Value ?? 1440)) * EmusPerTwip,
            1);

        // As faixas de capa e de página par vêm sempre: o Word guarda o `first`
        // mesmo com `w:titlePg` desligado, e é ele que volta quando se liga.
        BandDto? Band(bool header, HeaderFooterValues type)
        {
            // Da segunda seção em diante, ausente é herdada; declarada vazia é folha limpa.
            if (index > 0 && !Declares(section, header, type)) return null;
            var band = header
                ? HeaderReader.Read(section, part, inventory, type, contentWidthEmus)
                : HeaderReader.ReadFooter(section, part, inventory, type, contentWidthEmus);
            return index == 0 ? NullIfEmpty(band) : band;
        }

        return new PageSetupDto(
            Size: NearestSize(widthTwips, heightTwips, landscape),
            Orientation: landscape ? "landscape" : "portrait",
            Margins: new MarginsDto(
                Top: Millimeters(margin?.Top?.Value, 1440),
                Right: Millimeters((int?)margin?.Right?.Value, 1440),
                Bottom: Millimeters(margin?.Bottom?.Value, 1440),
                Left: Millimeters((int?)margin?.Left?.Value, 1440)),
            Header: Band(true, HeaderFooterValues.Default),
            Footer: Band(false, HeaderFooterValues.Default),
            FirstHeader: Band(true, HeaderFooterValues.First),
            FirstFooter: Band(false, HeaderFooterValues.First),
            EvenHeader: Band(true, HeaderFooterValues.Even),
            EvenFooter: Band(false, HeaderFooterValues.Even),
            HeaderDistanceMm: Millimeters((int?)margin?.Header?.Value, 708),
            FooterDistanceMm: Millimeters((int?)margin?.Footer?.Value, 708),
            PageNumberFormat: PageNumberFormatOf(section),
            PageNumberStart: StartElement(section.GetFirstChild<PageNumberType>()?.Start?.Value),
            TitlePage: HasTitlePage(section),
            EvenAndOddHeaders: UsesEvenAndOdd(part),
            Start: StartOf(section),
            Columns: ColumnsOf(section, inventory));
    }

    private static bool Declares(SectionProperties section, bool header, HeaderFooterValues type)
    {
        IEnumerable<HeaderFooterReferenceType> references = header
            ? section.Elements<HeaderReference>()
            : section.Elements<FooterReference>();
        return references.Any(reference =>
            (reference.Type?.Value ?? HeaderFooterValues.Default) == type &&
            !string.IsNullOrEmpty(reference.Id?.Value));
    }

    /// <summary>Com os padrões da especificação: uma coluna, 720 twips entre elas, sem linha.</summary>
    public static ColumnsDto ColumnsOf(SectionProperties section, Inventory? inventory = null)
    {
        var columns = section.GetFirstChild<Columns>();
        var widths = columns?.Elements<Column>().Select(column => Millimeters((int?)ParseTwips(column.Width?.Value), 0))
            .ToList();
        var count = Math.Clamp((int?)columns?.ColumnCount?.Value ?? (widths?.Count > 0 ? widths.Count : 1), 1, 45);
        var unequal = columns?.EqualWidth is not null && !columns.EqualWidth.Value && widths is { Count: > 1 };
        if (unequal && count > 1)
        {
            inventory?.NoteInvisible("colunas de larguras diferentes (mostradas iguais; o arquivo as mantém)");
        }

        return new ColumnsDto(
            count,
            Millimeters((int?)ParseTwips(columns?.Space?.Value), 720),
            columns?.Separator?.Value ?? false,
            unequal ? widths : null);
    }

    private static int? ParseTwips(string? value) =>
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var twips)
            ? twips
            : null;

    public static readonly string[] SectionStarts = ["nextPage", "continuous", "evenPage", "oddPage", "nextColumn"];

    /// <summary>"nextPage" quando falta, o padrão da especificação.</summary>
    public static string StartOf(SectionProperties section)
    {
        var name = section.GetFirstChild<SectionType>()?.Val?.InnerText;
        return name is not null && SectionStarts.Contains(name) ? name : "nextPage";
    }

    public static System.Text.Json.JsonElement StartElement(int? start) =>
        System.Text.Json.JsonSerializer.SerializeToElement(start);

    /// <summary><c>false</c> quando o campo está ausente.</summary>
    public static bool TryStartOf(PageSetupDto page, out int? start)
    {
        start = null;
        switch (page.PageNumberStart.ValueKind)
        {
            case System.Text.Json.JsonValueKind.Number when page.PageNumberStart.TryGetInt32(out var value):
                start = value;
                return true;
            case System.Text.Json.JsonValueKind.Null:
                return true;
            default:
                return false;
        }
    }

    public static readonly string[] PageNumberFormats =
        ["decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter"];

    /// <summary>Formato que o editor não desenha vira decimal na tela e volta intacto ao arquivo.</summary>
    public static string PageNumberFormatOf(SectionProperties section)
    {
        var name = section.GetFirstChild<PageNumberType>()?.Format?.InnerText;
        return name is not null && PageNumberFormats.Contains(name) ? name : "decimal";
    }

    /// <inheritdoc cref="HasTitlePage(SectionProperties)"/>
    public static bool TitlePageOf(SectionProperties section) => HasTitlePage(section);

    /// <inheritdoc cref="UsesEvenAndOdd(MainDocumentPart)"/>
    public static bool EvenAndOddOf(MainDocumentPart part) => UsesEvenAndOdd(part);

    /// <summary>Presente sem <c>w:val</c> é ligado, como todo interruptor do OOXML.</summary>
    private static bool HasTitlePage(SectionProperties section)
    {
        var flag = section.GetFirstChild<TitlePage>();
        return flag is not null && (flag.Val?.Value ?? true);
    }

    /// <summary>Mora em <c>settings.xml</c>: no Word é escolha do documento inteiro.</summary>
    private static bool UsesEvenAndOdd(MainDocumentPart part)
    {
        var flag = part.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>();
        return flag is not null && (flag.Val?.Value ?? true);
    }

    /// <summary>O modelo distingue "não tem" de "tem e está vazia".</summary>
    private static BandDto? NullIfEmpty(BandDto band) => band.IsEmpty ? null : band;

    private static double Millimeters(int? twips, int fallback) =>
        Math.Round((twips ?? fallback) / TwipsPerMillimeter, 1);

    private static bool IsKnownPaper(double widthTwips, double heightTwips)
    {
        var shortSide = Math.Min(widthTwips, heightTwips);
        var longSide = Math.Max(widthTwips, heightTwips);

        return Papers.Any(paper =>
            Math.Abs(shortSide - paper.Short) <= PaperTolerance
            && Math.Abs(longSide - paper.Long) <= PaperTolerance);
    }

    /// <summary>Pelo lado que atravessa a folha, o que separa A4 de Carta; empate fica com o primeiro.</summary>
    private static string NearestSize(double widthTwips, double heightTwips, bool landscape)
    {
        var across = landscape ? heightTwips : widthTwips;
        return Papers.MinBy(paper => Math.Abs(across - paper.Short)).Name;
    }

    /// <summary>
    /// Para não regravar o <c>w:sectPr</c> de quem não mexeu na página: um A5 viraria A4
    /// por uma correção de vírgula. Nas unidades do modelo, que dão a resolução da mudança.
    /// </summary>
    public static bool Matches(SectionProperties section, PageSetupDto page)
    {
        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        var margin = section.GetFirstChild<PageMargin>();

        var landscape = size?.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        // Sem `w:pgSz`, as medidas do primeiro papel da tabela.
        var widthTwips = (double?)size?.Width?.Value ?? Papers[0].Short;
        var heightTwips = (double?)size?.Height?.Value ?? Papers[0].Long;

        return landscape == string.Equals(page.Orientation, "landscape", StringComparison.Ordinal)
               && string.Equals(NearestSize(widthTwips, heightTwips, landscape), page.Size, StringComparison.Ordinal)
               && Millimeters(margin?.Top?.Value, 1440) == page.Margins.Top
               && Millimeters((int?)margin?.Right?.Value, 1440) == page.Margins.Right
               && Millimeters(margin?.Bottom?.Value, 1440) == page.Margins.Bottom
               && Millimeters((int?)margin?.Left?.Value, 1440) == page.Margins.Left;
    }

    /// <summary>Para a gravação preservar as medidas do arquivo enquanto são o papel que o modelo diz.</summary>
    public static string NameOfPaper(uint? widthTwips, uint? heightTwips, bool landscape) =>
        NearestSize(widthTwips ?? Papers[0].Short, heightTwips ?? Papers[0].Long, landscape);

    private static PageSetupDto Default() => new(
        "A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null,
        PageNumberFormat: "decimal", TitlePage: false, EvenAndOddHeaders: false);
}

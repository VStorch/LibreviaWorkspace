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
    // New, optional fields, so a `.sdoc` saved without them keeps opening.
    [property: JsonPropertyName("firstHeaderBand")] BandDto? FirstHeader = null,
    [property: JsonPropertyName("firstFooterBand")] BandDto? FirstFooter = null,
    [property: JsonPropertyName("evenHeaderBand")] BandDto? EvenHeader = null,
    [property: JsonPropertyName("evenFooterBand")] BandDto? EvenFooter = null,
    /// <c>w:pgMar/@header</c> and <c>@footer</c>: the vertical origin for anchors inside the band.
    [property: JsonPropertyName("headerDistanceMm")] double HeaderDistanceMm = 12.5,
    [property: JsonPropertyName("footerDistanceMm")] double FooterDistanceMm = 12.5,
    // The new document's plain text band, with `{n}` and `{total}` (PlainBandWriter).
    [property: JsonPropertyName("header")] string? HeaderText = null,
    [property: JsonPropertyName("footer")] string? FooterText = null,
    // `w:pgNumType` and the band switches: absent means "leave alone", not "turn off".
    [property: JsonPropertyName("pageNumberFormat")] string? PageNumberFormat = null,
    // Absent means "leave `w:start` alone"; null means "no start". `JsonElement` tells them apart.
    [property: JsonPropertyName("pageNumberStart")] System.Text.Json.JsonElement PageNumberStart = default,
    [property: JsonPropertyName("titlePage")] bool? TitlePage = null,
    [property: JsonPropertyName("evenAndOddHeaders")] bool? EvenAndOddHeaders = null,
    // `Id` only on sections before the last: the `sectionBreak` of the paragraph closing them.
    [property: JsonPropertyName("id")] string? Id = null,
    // `w:sectPr/w:type`. Absent means "leave alone".
    [property: JsonPropertyName("start")] string? Start = null,
    // `w:cols`. Absent means "leave alone".
    [property: JsonPropertyName("columns")] ColumnsDto? Columns = null);

/// <c>w:cols</c>. Unequal widths come in <c>WidthsMm</c>: the screen draws them equal, and the file
/// keeps them as long as the column count does not change.
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

/// <c>w:sectPr</c> → page setup.
public static class PageReader
{
    private const double TwipsPerMillimeter = Unit.TwipsPerInch / Unit.MillimetersPerInch;

    private const double EmusPerTwip = (double)Unit.EmusPerInch / Unit.TwipsPerInch;

    private const int DefaultMarginTwips = Unit.TwipsPerInch;

    /// <summary>
    /// Half a millimetre: the same A4 from LibreOffice and Word differs in the last twip.
    /// </summary>
    private const int PaperTolerance = 30;

    /// <summary>In twips, in a single table, so A4 does not have two widths.</summary>
    private static readonly (string Name, uint Short, uint Long)[] Papers =
    [
        ("A4", 11906U, 16838U),
        ("Letter", 12240U, 15840U),
    ];

    /// <summary>A4 for a name the table does not have.</summary>
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
    /// The last section is the model's "page"; the earlier ones come in order, with the id the
    /// reader puts on the paragraph closing them (<see cref="SectionIds"/>). Each carries only the
    /// bands it **declares**: inheritance belongs to whoever draws.
    /// </summary>
    public static (PageSetupDto Page, List<PageSetupDto>? Sections) ReadAll(
        Body body,
        MainDocumentPart part,
        Inventory inventory)
    {
        var all = body.Descendants<SectionProperties>().ToList();
        if (all.Count == 0) return (Default(), null);

        // A document without `w:sectPr` in the body is invalid, but it happens: the last paragraph
        // one stands in.
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

    /// <c>s1</c>, <c>s2</c>…, positional like the <c>oid</c>: the reference reading reproduces the
    /// ids.
    public static Dictionary<SectionProperties, string> SectionIds(Body body)
    {
        var ids = new Dictionary<SectionProperties, string>(ReferenceEqualityComparer.Instance);
        var all = body.Descendants<SectionProperties>().ToList();
        for (var index = 0; index < all.Count - 1; index++) ids[all[index]] = $"s{index + 1}";
        return ids;
    }

    private static PageSetupDto ReadOne(SectionProperties section, MainDocumentPart part, Inventory inventory, int index)
    {
        // The screen shows decimal and the file keeps asking for its own: a warning goes out.
        if (section.GetFirstChild<PageNumberType>()?.Format?.InnerText is { } pageFormat &&
            !PageNumberFormats.Contains(pageFormat))
        {
            inventory.NoteInvisible(
                $"formato de número de página \"{pageFormat}\" (mostrado em algarismos; o arquivo o mantém)");
        }
        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        var margin = section.GetFirstChild<PageMargin>();

        var landscape = size?.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        // Without `w:pgSz`, the measures of the table's first paper.
        var widthTwips = (double?)size?.Width?.Value ?? Papers[0].Short;
        var heightTwips = (double?)size?.Height?.Value ?? Papers[0].Long;

        // A paper other than A4 and Letter shows as the nearest, and the file keeps the measure:
        // invisibility.
        if (!IsKnownPaper(widthTwips, heightTwips))
        {
            inventory.NoteInvisible(
                "o tamanho do papel deste documento não é A4 nem Carta (ele é preservado no arquivo)");
        }

        // The band decides each piece's third against the column width.
        var contentWidthEmus = Math.Max(
            (widthTwips - (margin?.Left?.Value ?? DefaultMarginTwips) - (margin?.Right?.Value ?? DefaultMarginTwips)) * EmusPerTwip,
            1);

        // Title page and even page bands always come: Word keeps `first` even with `w:titlePg` off,
        // and that is what comes back when it is turned on.
        BandDto? Band(bool header, HeaderFooterValues type)
        {
            // From the second section on, absent means inherited; declared empty means a blank
            // sheet.
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
                Top: Millimeters(margin?.Top?.Value, DefaultMarginTwips),
                Right: Millimeters((int?)margin?.Right?.Value, DefaultMarginTwips),
                Bottom: Millimeters(margin?.Bottom?.Value, DefaultMarginTwips),
                Left: Millimeters((int?)margin?.Left?.Value, DefaultMarginTwips)),
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

    /// <summary>
    /// With the specification defaults: one column, 720 twips between them, no line.
    /// </summary>
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

    /// <summary>"nextPage" when missing, the specification default.</summary>
    public static string StartOf(SectionProperties section)
    {
        var name = section.GetFirstChild<SectionType>()?.Val?.InnerText;
        return name is not null && SectionStarts.Contains(name) ? name : "nextPage";
    }

    public static System.Text.Json.JsonElement StartElement(int? start) =>
        System.Text.Json.JsonSerializer.SerializeToElement(start);

    /// <c>false</c> when the field is absent.
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

    /// <summary>
    /// A format the editor does not draw becomes decimal on screen and goes back intact to the
    /// file.
    /// </summary>
    public static string PageNumberFormatOf(SectionProperties section)
    {
        var name = section.GetFirstChild<PageNumberType>()?.Format?.InnerText;
        return name is not null && PageNumberFormats.Contains(name) ? name : "decimal";
    }

    /// <inheritdoc cref="HasTitlePage(SectionProperties)"/>
    public static bool TitlePageOf(SectionProperties section) => HasTitlePage(section);

    /// <inheritdoc cref="UsesEvenAndOdd(MainDocumentPart)"/>
    public static bool EvenAndOddOf(MainDocumentPart part) => UsesEvenAndOdd(part);

    /// <summary>Present without <c>w:val</c> means on, like every OOXML switch.</summary>
    private static bool HasTitlePage(SectionProperties section)
    {
        var flag = section.GetFirstChild<TitlePage>();
        return flag is not null && (flag.Val?.Value ?? true);
    }

    /// <summary>It lives in <c>settings.xml</c>: in Word it is a document-wide choice.</summary>
    private static bool UsesEvenAndOdd(MainDocumentPart part)
    {
        var flag = part.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>();
        return flag is not null && (flag.Val?.Value ?? true);
    }

    /// <summary>The model tells "does not have" from "has and is empty".</summary>
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

    /// <summary>
    /// By the side crossing the sheet, which separates A4 from Letter; a tie goes to the first.
    /// </summary>
    private static string NearestSize(double widthTwips, double heightTwips, bool landscape)
    {
        var across = landscape ? heightTwips : widthTwips;
        return Papers.MinBy(paper => Math.Abs(across - paper.Short)).Name;
    }

    /// <summary>
    /// So the <c>w:sectPr</c> of someone who did not touch the page is not rewritten: an A5 would
    /// become A4 over a comma fix. In model units, which give the change its resolution.
    /// </summary>
    public static bool Matches(SectionProperties section, PageSetupDto page)
    {
        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        var margin = section.GetFirstChild<PageMargin>();

        var landscape = size?.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        // Without `w:pgSz`, the measures of the table's first paper.
        var widthTwips = (double?)size?.Width?.Value ?? Papers[0].Short;
        var heightTwips = (double?)size?.Height?.Value ?? Papers[0].Long;

        return landscape == string.Equals(page.Orientation, "landscape", StringComparison.Ordinal)
               && string.Equals(NearestSize(widthTwips, heightTwips, landscape), page.Size, StringComparison.Ordinal)
               && Millimeters(margin?.Top?.Value, DefaultMarginTwips) == page.Margins.Top
               && Millimeters((int?)margin?.Right?.Value, DefaultMarginTwips) == page.Margins.Right
               && Millimeters(margin?.Bottom?.Value, DefaultMarginTwips) == page.Margins.Bottom
               && Millimeters((int?)margin?.Left?.Value, DefaultMarginTwips) == page.Margins.Left;
    }

    /// <summary>
    /// So saving keeps the file's measures as long as they are the paper the model names.
    /// </summary>
    public static string NameOfPaper(uint? widthTwips, uint? heightTwips, bool landscape) =>
        NearestSize(widthTwips ?? Papers[0].Short, heightTwips ?? Papers[0].Long, landscape);

    private static PageSetupDto Default() => new(
        "A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null,
        PageNumberFormat: "decimal", TitlePage: false, EvenAndOddHeaders: false);
}

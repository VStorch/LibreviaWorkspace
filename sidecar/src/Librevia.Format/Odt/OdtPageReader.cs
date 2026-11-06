using System.Text;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The first block's master page → the page setup: paper, orientation and margins. ODF measures the
/// top margin to the header and Word to the text, so the header's height and spacing join the
/// margin, the inverse of <see cref="OdtTemplatePage"/>. Only a plain text header or footer has a place in
/// the model (<c>{n}</c> and <c>{total}</c> for the page fields).
/// </summary>
internal static class OdtPageReader
{
    public const string DefaultMaster = "Standard";

    /// <summary>Half a millimetre each way: LibreOffice's A4 is 21.001 cm.</summary>
    private const double PaperToleranceMm = 1;

    private const double DefaultMarginMm = 20;
    private const double DefaultBandDistanceMm = 12.5;

    private static readonly (string Name, double Short, double Long)[] Papers =
    [
        ("A4", 210, 297),
        ("Letter", 215.9, 279.4),
    ];

    private const string Centered = "center";

    private static readonly string[] Variants = ["header-left", "header-first", "footer-left", "footer-first"];

    public static PageSetupDto Read(XDocument? styles, string? masterName, Inventory inventory)
    {
        var masters = styles?.Root?.Element(Ns.Office + "master-styles")?.Elements(Ns.Style + "master-page").ToList() ?? [];
        var master = masters.FirstOrDefault(page => page.Attr(Ns.Style + "name") == (masterName ?? DefaultMaster))
                     ?? masters.FirstOrDefault();
        var layoutName = master?.Attr(Ns.Style + "page-layout-name");
        var layout = styles?.Root?.Element(Ns.Office + "automatic-styles")?.Elements(Ns.Style + "page-layout")
            .FirstOrDefault(element => element.Attr(Ns.Style + "name") == layoutName);
        var properties = layout?.Element(Ns.Style + "page-layout-properties");

        var (size, orientation) = PaperOf(properties, inventory);
        var bandStyles = styles is null ? null : OdtStyles.Read(styles, styles);
        var header = BandText(master, "header", bandStyles, inventory);
        var footer = BandText(master, "footer", bandStyles, inventory);
        var (top, headerDistance) = Edge(properties?.Attr(Ns.Fo + "margin-top"), layout, "header-style", "margin-bottom", header);
        var (bottom, footerDistance) = Edge(properties?.Attr(Ns.Fo + "margin-bottom"), layout, "footer-style", "margin-top", footer);
        if (master is not null && Variants.Any(variant => HasContent(master.Element(Ns.Style + variant))))
        {
            inventory.NoteLoss(OdtNotices.HeaderVariants);
        }

        var margins = new MarginsDto(
            top,
            OdtLength.RoundedMillimetres(properties?.Attr(Ns.Fo + "margin-right")) ?? DefaultMarginMm,
            bottom,
            OdtLength.RoundedMillimetres(properties?.Attr(Ns.Fo + "margin-left")) ?? DefaultMarginMm);
        return new PageSetupDto(size, orientation, margins, Header: null, Footer: null,
            HeaderDistanceMm: headerDistance, FooterDistanceMm: footerDistance,
            HeaderText: header ?? string.Empty, FooterText: footer ?? string.Empty,
            PageNumberFormat: PageNumberFormatOf(properties?.Attr(Ns.Style + "num-format")),
            PageNumberStart: PageReader.StartElement(null),
            TitlePage: false,
            EvenAndOddHeaders: false);
    }

    /// <summary>The five the panel offers; any other is shown as digits.</summary>
    private static string PageNumberFormatOf(string? format) =>
        OdtListLevels.FormatOf(format) is { } known and not "none" ? known : "decimal";

    private static (string Size, string Orientation) PaperOf(XElement? properties, Inventory inventory)
    {
        var width = OdtLength.Millimetres(properties?.Attr(Ns.Fo + "page-width"));
        var height = OdtLength.Millimetres(properties?.Attr(Ns.Fo + "page-height"));
        if (width is not { } across || height is not { } down) return (Papers[0].Name, "portrait");

        var orientation = properties?.Attr(Ns.Style + "print-orientation") ?? (across > down ? "landscape" : "portrait");
        var (shortSide, longSide) = (Math.Min(across, down), Math.Max(across, down));
        foreach (var paper in Papers)
        {
            if (Math.Abs(paper.Short - shortSide) <= PaperToleranceMm && Math.Abs(paper.Long - longSide) <= PaperToleranceMm)
            {
                return (paper.Name, orientation);
            }
        }

        inventory.NoteLoss(OdtNotices.PaperSize);
        return (Papers[0].Name, orientation);
    }

    /// <summary>The margin to the text, and the distance from the edge to the band.</summary>
    private static (double Margin, double Distance) Edge(
        string? pageMargin, XElement? layout, string bandStyle, string spacing, string? band)
    {
        var edge = OdtLength.Millimetres(pageMargin) ?? DefaultMarginMm;
        if (band is null) return (Math.Round(edge, 2), DefaultBandDistanceMm);

        var properties = layout?.Element(Ns.Style + bandStyle)?.Element(Ns.Style + "header-footer-properties");
        var height = OdtLength.Millimetres(properties?.Attr(Ns.Fo + "min-height") ?? properties?.Attr(Ns.Svg + "height")) ?? 0;
        var gap = OdtLength.Millimetres(properties?.Attr(Ns.Fo + spacing)) ?? 0;
        return (Math.Round(edge + height + gap, 2), Math.Round(edge, 2));
    }

    /// <summary>
    /// The band as one line of text, or <c>null</c> when the page shows none. The plain band is
    /// centered, in the editor's own look: any other alignment or character formatting is noted.
    /// </summary>
    private static string? BandText(XElement? master, string kind, OdtStyles? styles, Inventory inventory)
    {
        var band = master?.Element(Ns.Style + kind);
        if (band is null || !HasContent(band)) return null;

        var lines = new List<string>();
        foreach (var block in band.Elements())
        {
            if (block.Name != Ns.Text + "p" && block.Name != Ns.Text + "h")
            {
                inventory.NoteLoss(OdtNotices.BandFormatting);
                continue;
            }

            if (HasOwnLook(block, styles)) inventory.NoteLoss(OdtNotices.BandFormatting);

            var text = new StringBuilder();
            AppendBand(block, text, inventory);
            var line = CollapseSpaces(text.ToString());
            if (line.Length > 0) lines.Add(line);
        }

        if (lines.Count > 1) inventory.NoteLoss(OdtNotices.BandFormatting);
        return string.Join(' ', lines);
    }

    /// <summary>
    /// Centered by alignment, or by a center tab stop as older Librevia exports wrote the plain band, and with
    /// no character formatting beyond the plain band's own look.
    /// </summary>
    private static bool HasOwnLook(XElement block, OdtStyles? styles)
    {
        if (styles is null) return true;

        var chain = styles.Chain(OdtStyles.ParagraphFamily, block.Attr(Ns.Text + "style-name"));
        var centered = styles.Paragraph(chain).TextAlign == Centered || IsCenteredByTab(block, chain);
        if (!centered || IsOwnCharacterLook(styles.Character(chain, CharacterFormat.Empty))) return true;

        return block.Descendants(Ns.Text + "span").Any(span => IsOwnCharacterLook(
            styles.Character(chain.Concat(styles.Chain(OdtStyles.TextFamily, span.Attr(Ns.Text + "style-name"))), CharacterFormat.Empty)));
    }

    /// <summary>A single leading tab to a center stop.</summary>
    private static bool IsCenteredByTab(XElement block, List<OdtStyle> chain)
    {
        var leading = block.Nodes().FirstOrDefault(node => node is not XText { Value.Length: 0 });
        var hasCenterStop = chain.Any(style => style.Properties("paragraph-properties")?
            .Descendants(Ns.Style + "tab-stop").Any(stop => stop.Attr(Ns.Style + "type") == Centered) == true);
        return hasCenterStop
               && leading is XElement { Name.LocalName: "tab" } tab && tab.Name.Namespace == Ns.Text
               && block.Descendants(Ns.Text + "tab").Count() == 1;
    }

    private static bool IsOwnCharacterLook(CharacterFormat format) =>
        format with
        {
            FontFamily = format.FontFamily?.StartsWith(BuiltinStyles.BandFont, StringComparison.Ordinal) == true
                ? null
                : format.FontFamily,
            FontSizePt = format.FontSizePt == PlainBandWriter.BandSizePt ? null : format.FontSizePt,
            Color = string.Equals(format.Color, "#" + PlainBandWriter.BandColor, StringComparison.OrdinalIgnoreCase)
                ? null
                : format.Color,
            Bold = OnlyWhenOn(format.Bold),
            Italic = OnlyWhenOn(format.Italic),
            Underline = OnlyWhenOn(format.Underline),
            Strike = OnlyWhenOn(format.Strike),
            Caps = OnlyWhenOn(format.Caps),
            SmallCaps = OnlyWhenOn(format.SmallCaps),
            Position = format.Position == CharacterFormat.Baseline ? null : format.Position,
            Background = format.Background == CharacterFormat.Transparent ? null : format.Background,
        } != CharacterFormat.Empty;

    private static bool? OnlyWhenOn(bool? value) => value == true ? true : null;

    private static bool HasContent(XElement? band) =>
        band is not null && band.Attr(Ns.Style + "display") != "false" && band.HasElements;

    private static void AppendBand(XElement container, StringBuilder text, Inventory inventory)
    {
        foreach (var node in container.Nodes())
        {
            if (node is XText plain) text.Append(plain.Value);
            else if (node is XElement element) AppendBandElement(element, text, inventory);
        }
    }

    private static void AppendBandElement(XElement element, StringBuilder text, Inventory inventory)
    {
        switch (element.Name.LocalName)
        {
            case "page-number" when element.Name.Namespace == Ns.Text:
                text.Append("{n}");
                return;
            case "page-count" when element.Name.Namespace == Ns.Text:
                text.Append("{total}");
                return;
            case "s" or "tab" or "line-break" when element.Name.Namespace == Ns.Text:
                text.Append(' ');
                return;
            case "span" or "a" when element.Name.Namespace == Ns.Text:
                AppendBand(element, text, inventory);
                return;
        }

        if (OdtNotices.IsTextField(element))
        {
            inventory.NoteLoss(Inventory.HeaderFields);
            AppendBand(element, text, inventory);
            return;
        }

        inventory.NoteLoss(OdtNotices.BandFormatting);
    }

    /// <summary>ODF collapses white space; the tab stops that center the band become nothing.</summary>
    private static string CollapseSpaces(string text) =>
        string.Join(' ', text.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries));
}

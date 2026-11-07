using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A section's page layout and master page. ODF measures the margin to the header, the model to the
/// text: with a header, the top margin is the header's distance, and its minimum height covers the
/// rest, so a taller header pushes the text, as in Word. The inverse of <see cref="OdtPageReader"/>.
/// </summary>
internal static class OdtTemplatePage
{
    private const string Landscape = "landscape";

    public static XElement Layout(string name, PageSetupDto page)
    {
        var portrait = page.Orientation != Landscape;
        var (shortSide, longSide) = OdtStylesPart.Papers.GetValueOrDefault(page.Size, OdtStylesPart.Papers["A4"]);
        var header = HasAny(page, page.Header, page.HeaderText, page.FirstHeader, page.EvenHeader);
        var footer = HasAny(page, page.Footer, page.FooterText, page.FirstFooter, page.EvenFooter);
        var top = header ? Math.Min(page.HeaderDistanceMm, page.Margins.Top) : page.Margins.Top;
        var bottom = footer ? Math.Min(page.FooterDistanceMm, page.Margins.Bottom) : page.Margins.Bottom;

        var properties = new XElement(
            Ns.Style + "page-layout-properties",
            new XAttribute(Ns.Fo + "page-width", OdtFormatWriter.Mm(portrait ? shortSide : longSide)!),
            new XAttribute(Ns.Fo + "page-height", OdtFormatWriter.Mm(portrait ? longSide : shortSide)!),
            new XAttribute(Ns.Style + "print-orientation", portrait ? "portrait" : Landscape),
            new XAttribute(Ns.Fo + "margin-top", OdtFormatWriter.Mm(Math.Max(0, top))!),
            new XAttribute(Ns.Fo + "margin-bottom", OdtFormatWriter.Mm(Math.Max(0, bottom))!),
            new XAttribute(Ns.Fo + "margin-left", OdtFormatWriter.Mm(page.Margins.Left)!),
            new XAttribute(Ns.Fo + "margin-right", OdtFormatWriter.Mm(page.Margins.Right)!),
            new XAttribute(Ns.Style + "num-format", OdtStylesPart.PageNumberFormats.GetValueOrDefault(page.PageNumberFormat ?? "decimal", "1")),
            new XAttribute(Ns.Style + "writing-mode", "lr-tb"));
        return new XElement(
            Ns.Style + "page-layout",
            new XAttribute(Ns.Style + "name", name),
            properties,
            BandStyle("header", header, page.Margins.Top - top, "margin-bottom"),
            BandStyle("footer", footer, page.Margins.Bottom - bottom, "margin-top"));
    }

    private static XElement BandStyle(string kind, bool present, double heightMm, string spacing)
    {
        var style = new XElement(Ns.Style + kind + "-style");
        if (!present) return style;
        style.Add(new XElement(
            Ns.Style + "header-footer-properties",
            new XAttribute(Ns.Fo + "min-height", OdtFormatWriter.Mm(Math.Max(0, heightMm))!),
            new XAttribute(Ns.Fo + spacing, "0mm"),
            new XAttribute(Ns.Style + "dynamic-spacing", "false")));
        return style;
    }

    public static XElement Master(string name, string layout, PageSetupDto page, OdtBandWriter bands, OdtWriteContext context)
    {
        var master = new XElement(
            Ns.Style + "master-page",
            new XAttribute(Ns.Style + "name", name),
            new XAttribute(Ns.Style + "page-layout-name", layout));
        var width = ContentWidthMm(page);
        master.Add(
            Band("header", page.Header, page.HeaderText, width, bands, context),
            page.EvenAndOddHeaders == true ? Band("header-left", page.EvenHeader, null, width, bands, context) : null,
            page.TitlePage == true ? Band("header-first", page.FirstHeader, null, width, bands, context) : null,
            Band("footer", page.Footer, page.FooterText, width, bands, context),
            page.EvenAndOddHeaders == true ? Band("footer-left", page.EvenFooter, null, width, bands, context) : null,
            page.TitlePage == true ? Band("footer-first", page.FirstFooter, null, width, bands, context) : null);
        return master;
    }

    public static double ContentWidthMm(PageSetupDto page)
    {
        var (shortSide, longSide) = OdtStylesPart.Papers.GetValueOrDefault(page.Size, OdtStylesPart.Papers["A4"]);
        var width = page.Orientation == Landscape ? longSide : shortSide;
        return width - page.Margins.Left - page.Margins.Right;
    }

    /// <summary>A band from another format as it is; otherwise the plain text band, as saving writes it.</summary>
    private static XElement? Band(string kind, BandDto? band, string? text, double widthMm, OdtBandWriter bands, OdtWriteContext context)
    {
        if (band is not null) return HasContent(band) ? new XElement(Ns.Style + kind, bands.Write(band, widthMm, kind.StartsWith("header", StringComparison.Ordinal))) : null;
        return string.IsNullOrWhiteSpace(text) ? null : OdtStylesPart.PlainBand(kind, text, null, context.Automatic, context.Styles.Fonts);
    }

    private static bool HasContent(BandDto? band) => band is not null && !band.IsEmpty;

    private static bool HasAny(PageSetupDto page, BandDto? main, string? text, BandDto? first, BandDto? even) =>
        (main is not null ? HasContent(main) : !string.IsNullOrWhiteSpace(text)) ||
        (page.TitlePage == true && HasContent(first)) ||
        (page.EvenAndOddHeaders == true && HasContent(even));
}

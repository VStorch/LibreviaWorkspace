using System.Text;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>styles.xml</c>, only where the model differs from the file: a named style changed or created,
/// the page layout of each master page a section uses (or a copy, see <see cref="OdtSectionMasters"/>),
/// and its header or footer when their text changed.
/// Each touched element is cut out and replaced; the rest of the part stays as it was.
/// </summary>
internal static class OdtStylesPart
{
    private const double MarginToleranceMm = 0.01;
    private const string HeaderStyle = "Header";
    private const string FooterStyle = "Footer";

    internal static readonly Dictionary<string, (double Short, double Long)> Papers = new(StringComparer.Ordinal)
    {
        ["A4"] = (210, 297),
        ["Letter"] = (215.9, 279.4),
    };

    internal static readonly Dictionary<string, string> PageNumberFormats = new(StringComparer.Ordinal)
    {
        ["decimal"] = "1",
        ["lowerRoman"] = "i",
        ["upperRoman"] = "I",
        ["lowerLetter"] = "a",
        ["upperLetter"] = "A",
    };

    /// <returns>The new text, or <c>null</c> when nothing changed.</returns>
    public static string? Write(OdtSource source, DocumentModelDto model, IEnumerable<string> taken, IReadOnlyList<OdtMasterChange> masters)
    {
        if (source.Package.Styles?.Root is not { } root) return null;

        var xml = OdtWriter.TextOf(source.Package, OdtPackage.StylesPath);
        var spans = OdtXmlSpans.Root(xml);
        var xmlOut = new OdtXmlOut(root);
        var edits = new TextEdits(xml);
        var changed = NamedStyles(source, model.Styles, root, spans, xmlOut, edits);

        var automatic = new OdtAutomaticStyles(taken);
        var copies = new List<XElement>();
        foreach (var master in masters) changed |= Page(source, master, new PartEdits(root, spans, xmlOut, edits, automatic), copies);
        if (copies.Count > 0 && root.Element(Ns.Office + "master-styles") is { } masterStyles)
        {
            edits.AppendContent(OdtXmlSpans.Find(spans, masterStyles), xmlOut.Write(copies));
        }

        OdtWriter.AppendStyles(edits, spans, root, xmlOut, automatic.Created, root.Element(Ns.Office + "master-styles"));
        return changed ? edits.Apply() : null;
    }

    private static bool NamedStyles(
        OdtSource source, StyleSheetDto? sheet, XElement root, XmlSpan spans, OdtXmlOut xmlOut, TextEdits edits)
    {
        if (sheet is null) return false;
        var office = root.Element(Ns.Office + "styles");
        var created = new List<XElement>();
        var changed = false;
        foreach (var (id, wanted) in sheet.Styles)
        {
            var known = source.Sheet.Styles.GetValueOrDefault(id);
            if (known is not null && Same(known, wanted)) continue;
            if (wanted.Type is not ("paragraph" or "character")) continue;

            var family = wanted.Type == "paragraph" ? OdtStyles.ParagraphFamily : OdtStyles.TextFamily;
            var element = known is null ? null : source.Styles.Find(family, OdtWordStyles.OdfNameOf(family, id))?.Element;
            if (element is { Parent: { } parent } && parent == office)
            {
                edits.Replace(OdtXmlSpans.Find(spans, element), xmlOut.Write(Changed(source, element, known!, wanted, family)));
                changed = true;
            }
            else if (known is null)
            {
                created.Add(Created(source, id, wanted, family));
            }
        }

        if (created.Count == 0 || office is null) return changed;
        edits.AppendContent(OdtXmlSpans.Find(spans, office), xmlOut.Write(created));
        return true;
    }

    /// <summary>What the reader reads and the editor edits; the panel's own flags do not count.</summary>
    private static bool Same(StyleDefinitionDto known, StyleDefinitionDto wanted) =>
        known.Name == wanted.Name && known.BasedOn == wanted.BasedOn && known.Next == wanted.Next &&
        (known.Paragraph ?? new StyleParagraphDto()) == (wanted.Paragraph ?? new StyleParagraphDto()) &&
        (known.Character ?? new StyleCharacterDto()) == (wanted.Character ?? new StyleCharacterDto());

    /// <summary>The file's element with only the changed fields rewritten: tabs, borders and the rest stay.</summary>
    private static XElement Changed(OdtSource source, XElement original, StyleDefinitionDto known, StyleDefinitionDto wanted, string family)
    {
        var element = new XElement(original);
        if (known.Name != wanted.Name && OdtWordStyles.NameOf(family, original.Attr(Ns.Style + "name") ?? string.Empty) is null)
        {
            element.SetAttributeValue(Ns.Style + "display-name", wanted.Name);
        }

        if (known.BasedOn != wanted.BasedOn) element.SetAttributeValue(Ns.Style + "parent-style-name", OdfName(source, family, wanted.BasedOn));
        if (known.Next != wanted.Next) element.SetAttributeValue(Ns.Style + "next-style-name", OdfName(source, family, wanted.Next));

        if (family == OdtStyles.ParagraphFamily)
        {
            OdtStyleFields.Paragraph(
                Properties(element, "paragraph-properties", before: "text-properties"),
                OdtFormatWriter.OfStyle(known.Paragraph),
                OdtFormatWriter.OfStyle(wanted.Paragraph));
        }

        OdtStyleFields.Text(
            Properties(element, "text-properties", before: null),
            OdtFormatWriter.OfStyle(known.Character),
            OdtFormatWriter.OfStyle(wanted.Character),
            source.Styles.Fonts);
        foreach (var empty in element.Elements().Where(child => !child.HasAttributes && !child.HasElements).ToList()) empty.Remove();
        return element;
    }

    private static XElement Properties(XElement style, string kind, string? before)
    {
        if (style.Element(Ns.Style + kind) is { } existing) return existing;
        var created = new XElement(Ns.Style + kind);
        if (before is not null && style.Element(Ns.Style + before) is { } next) next.AddBeforeSelf(created);
        else style.Add(created);
        return created;
    }

    private static XElement Created(OdtSource source, string id, StyleDefinitionDto wanted, string family)
    {
        var element = new XElement(
            Ns.Style + "style",
            new XAttribute(Ns.Style + "name", SafeName(id)),
            new XAttribute(Ns.Style + "display-name", wanted.Name),
            new XAttribute(Ns.Style + "family", family));
        if (OdfName(source, family, wanted.BasedOn) is { } parent) element.Add(new XAttribute(Ns.Style + "parent-style-name", parent));
        if (OdfName(source, family, wanted.Next) is { } next) element.Add(new XAttribute(Ns.Style + "next-style-name", next));
        if (family == OdtStyles.ParagraphFamily) element.Add(OdtFormatWriter.Paragraph(OdtFormatWriter.OfStyle(wanted.Paragraph)));
        element.Add(OdtFormatWriter.Text(OdtFormatWriter.OfStyle(wanted.Character), source.Styles.Fonts));
        foreach (var empty in element.Elements().Where(child => !child.HasAttributes).ToList()) empty.Remove();
        return element;
    }

    private static string? OdfName(OdtSource source, string family, string? id)
    {
        if (string.IsNullOrEmpty(id)) return null;
        var name = OdtWordStyles.OdfNameOf(family, id);
        return source.Styles.IsNamed(family, name) ? name : SafeName(id);
    }

    /// <summary>A <c>style:name</c> is an XML name: LibreOffice escapes the rest as <c>_20_</c>.</summary>
    internal static string SafeName(string id)
    {
        var name = new StringBuilder();
        foreach (var letter in id)
        {
            var allowed = char.IsLetter(letter) || letter == '_' || (name.Length > 0 && (char.IsDigit(letter) || letter is '-' or '.'));
            if (allowed) name.Append(letter);
            else name.Append('_').Append(((int)letter).ToString("x2", System.Globalization.CultureInfo.InvariantCulture)).Append('_');
        }

        return name.ToString();
    }

    private sealed record PartEdits(XElement Root, XmlSpan Spans, OdtXmlOut Out, TextEdits Edits, OdtAutomaticStyles Automatic);

    /// <summary>What <see cref="Page"/> writes: the layout and the plain bands.</summary>
    public static bool SameLayout(PageSetupDto known, PageSetupDto wanted) =>
        !BandsChanged(known, wanted) && known.Size == wanted.Size && known.Orientation == wanted.Orientation &&
        SameMargins(known.Margins, wanted.Margins) && known.PageNumberFormat == wanted.PageNumberFormat;

    private static bool BandsChanged(PageSetupDto known, PageSetupDto wanted) =>
        (wanted.Header is null && (known.HeaderText ?? string.Empty) != (wanted.HeaderText ?? string.Empty)) ||
        (wanted.Footer is null && (known.FooterText ?? string.Empty) != (wanted.FooterText ?? string.Empty));

    /// <summary>The master page brought to the section's setup, in place or, with <c>From</c>, as a new copy.</summary>
    private static bool Page(OdtSource source, OdtMasterChange change, PartEdits part, List<XElement> copies)
    {
        var (name, from, known, wanted) = change;
        var masters = part.Root.Element(Ns.Office + "master-styles")?.Elements(Ns.Style + "master-page").ToList() ?? [];
        var master = masters.FirstOrDefault(page => page.Attr(Ns.Style + "name") == (from ?? name))
                     ?? (from is null ? masters.FirstOrDefault() : null);
        if (master is null) return false;

        var bands = new XElement(master);
        var bandsChanged = BandsChanged(known, wanted);
        if (bandsChanged)
        {
            if (wanted.Header is null) Band(source, bands, "header", known.HeaderText, wanted.HeaderText, HeaderStyle, part.Automatic);
            if (wanted.Footer is null) Band(source, bands, "footer", known.FooterText, wanted.FooterText, FooterStyle, part.Automatic);
        }

        var layout = part.Root.Element(Ns.Office + "automatic-styles")?.Elements(Ns.Style + "page-layout")
            .FirstOrDefault(element => element.Attr(Ns.Style + "name") == master.Attr(Ns.Style + "page-layout-name"));
        var properties = layout?.Element(Ns.Style + "page-layout-properties");
        var changed = properties is null ? null : Layout(properties, layout!, bands, known, wanted);
        if (from is not null)
        {
            copies.Add(Copy(bands, name, layout, properties, changed, part.Automatic));
            return true;
        }

        if (bandsChanged) part.Edits.Replace(OdtXmlSpans.Find(part.Spans, master), part.Out.Write(bands));
        if (changed is not null) part.Edits.Replace(OdtXmlSpans.Find(part.Spans, properties!), part.Out.Write(changed));
        return true;
    }

    /// <summary>A new master page, with its own page layout when the layout changes.</summary>
    private static XElement Copy(
        XElement master, string name, XElement? layout, XElement? properties, XElement? changed, OdtAutomaticStyles automatic)
    {
        master.SetAttributeValue(Ns.Style + "name", name);
        master.SetAttributeValue(Ns.Style + "display-name", null);
        if (layout is null || changed is null) return master;
        var copy = new XElement(
            layout.Name,
            layout.Attributes().Where(attribute => attribute.Name != Ns.Style + "name"),
            layout.Elements().Select(element => element == properties ? changed : new XElement(element)));
        master.SetAttributeValue(Ns.Style + "page-layout-name", automatic.Add(OdtAutomaticStyles.PageLayout, copy));
        return master;
    }

    private static bool SameMargins(MarginsDto known, MarginsDto wanted) =>
        Math.Abs(known.Top - wanted.Top) < MarginToleranceMm && Math.Abs(known.Bottom - wanted.Bottom) < MarginToleranceMm &&
        Math.Abs(known.Left - wanted.Left) < MarginToleranceMm && Math.Abs(known.Right - wanted.Right) < MarginToleranceMm;

    /// <summary>
    /// The plain band of the model (<c>header</c>, <c>footer</c>), as one centered line in the
    /// editor's own look, the one <see cref="OdtPageReader"/> reads back as plain; <c>{n}</c> and
    /// <c>{total}</c> are page fields.
    /// </summary>
    internal static XElement PlainBand(string kind, string text, string? parent, OdtAutomaticStyles automatic, OdtFonts fonts)
    {
        var centered = new XElement(Ns.Style + "paragraph-properties", new XAttribute(Ns.Fo + "text-align", "center"));
        var look = OdtFormatWriter.Text(
            new CharacterFormat(FontFamily: BuiltinStyles.BandFont, FontSizePt: PlainBandWriter.BandSizePt, Color: "#" + PlainBandWriter.BandColor),
            fonts);
        var style = automatic.Style(OdtAutomaticStyles.Band, OdtStyles.ParagraphFamily, parent, centered, look);
        return new XElement(Ns.Style + kind, new XElement(Ns.Text + "p", new XAttribute(Ns.Text + "style-name", style), BandText(text)));
    }

    /// <summary>The band as <see cref="PlainBand"/> writes it; empty removes it.</summary>
    private static void Band(
        OdtSource source, XElement master, string kind, string? known, string? wanted, string namedStyle, OdtAutomaticStyles automatic)
    {
        if ((known ?? string.Empty) == (wanted ?? string.Empty)) return;
        var existing = master.Element(Ns.Style + kind);
        if (string.IsNullOrWhiteSpace(wanted))
        {
            existing?.Remove();
            return;
        }

        var parent = source.Styles.IsNamed(OdtStyles.ParagraphFamily, namedStyle) ? namedStyle : null;
        var band = PlainBand(kind, wanted, parent, automatic, source.Styles.Fonts);

        if (existing is not null) existing.ReplaceWith(band);
        else if (kind == "header") master.AddFirst(band);
        else if (master.Elements().LastOrDefault(child => child.Name.LocalName.StartsWith("header", StringComparison.Ordinal)) is { } header) header.AddAfterSelf(band);
        else master.AddFirst(band);
    }

    private static IEnumerable<XNode> BandText(string text)
    {
        var parts = text.Split(["{n}", "{total}"], StringSplitOptions.None);
        var fields = new List<string>();
        for (var index = 0; index < text.Length;)
        {
            var page = text.IndexOf("{n}", index, StringComparison.Ordinal);
            var total = text.IndexOf("{total}", index, StringComparison.Ordinal);
            if (page < 0 && total < 0) break;
            var isPage = page >= 0 && (total < 0 || page < total);
            fields.Add(isPage ? "page-number" : "page-count");
            index = (isPage ? page : total) + (isPage ? 3 : 7);
        }

        for (var index = 0; index < parts.Length; index++)
        {
            if (parts[index].Length > 0) yield return new XText(parts[index]);
            if (index < fields.Count)
            {
                yield return fields[index] == "page-number"
                    ? new XElement(Ns.Text + "page-number", new XAttribute(Ns.Text + "select-page", "current"), "1")
                    : new XElement(Ns.Text + "page-count", "1");
            }
        }
    }

    /// <summary>
    /// The page layout's properties for the model's page, or <c>null</c> when they already say it.
    /// ODF measures the top margin to the header, the model to the text: the header's height and
    /// spacing come out of the margin, the inverse of <see cref="OdtPageReader"/>.
    /// </summary>
    private static XElement? Layout(XElement original, XElement layout, XElement master, PageSetupDto known, PageSetupDto page)
    {
        var properties = new XElement(original);
        if (known.Size != page.Size || known.Orientation != page.Orientation)
        {
            var portrait = page.Orientation != "landscape";
            var (shortSide, longSide) = Papers.GetValueOrDefault(page.Size, Papers["A4"]);
            Length(properties, "page-width", portrait ? shortSide : longSide, original);
            Length(properties, "page-height", portrait ? longSide : shortSide, original);
            properties.SetAttributeValue(Ns.Style + "print-orientation", portrait ? "portrait" : "landscape");
        }

        Length(properties, "margin-left", page.Margins.Left, original);
        Length(properties, "margin-right", page.Margins.Right, original);
        Length(properties, "margin-top", page.Margins.Top - BandSpace(layout, master, "header", "margin-bottom"), original);
        Length(properties, "margin-bottom", page.Margins.Bottom - BandSpace(layout, master, "footer", "margin-top"), original);

        if (page.PageNumberFormat is { } format && format != known.PageNumberFormat)
        {
            properties.SetAttributeValue(Ns.Style + "num-format", PageNumberFormats.GetValueOrDefault(format, "1"));
        }

        return XNode.DeepEquals(properties, original) ? null : properties;
    }

    /// <summary>Only when the measure differs: the file's own unit stays otherwise.</summary>
    private static void Length(XElement properties, string name, double mm, XElement original)
    {
        if (OdtLength.Millimetres(original.Attr(Ns.Fo + name)) is { } current && Math.Abs(current - mm) < MarginToleranceMm) return;
        properties.SetAttributeValue(Ns.Fo + name, OdtFormatWriter.Mm(Math.Max(0, mm)));
    }

    private static double BandSpace(XElement layout, XElement master, string kind, string spacing)
    {
        var band = master.Element(Ns.Style + kind);
        if (band is null || band.Attr(Ns.Style + "display") == "false" || !band.HasElements) return 0;
        var properties = layout.Element(Ns.Style + kind + "-style")?.Element(Ns.Style + "header-footer-properties");
        var height = OdtLength.Millimetres(properties?.Attr(Ns.Fo + "min-height") ?? properties?.Attr(Ns.Svg + "height")) ?? 0;
        return height + (OdtLength.Millimetres(properties?.Attr(Ns.Fo + spacing)) ?? 0);
    }
}

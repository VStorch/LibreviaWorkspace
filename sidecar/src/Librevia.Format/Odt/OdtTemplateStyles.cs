using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The template's <c>styles.xml</c>: the defaults and the named styles of the model under the names
/// LibreOffice knows (<see cref="OdtWordStyles"/>), headings without numbers, the note numbering,
/// and a page layout and master page per section, with its headers and footers.
/// </summary>
internal static class OdtTemplateStyles
{
    private const string TabStopDistance = "12.7mm";
    private const int OutlineLevels = 10;
    private const int WordOutlineLevels = 9;
    private const string LayoutPrefix = "LvPm";

    private static readonly Dictionary<string, string> NoteRestarts = new(StringComparer.Ordinal)
    {
        ["eachPage"] = "page",
        ["eachSect"] = "chapter",
    };

    public static XDocument Write(OdtCreateDto request, OdtPictures pictures)
    {
        var root = OdtTemplate.Root(Ns.Office + "document-styles");
        var office = new XElement(Ns.Office + "styles", Defaults(request.Styles), NamedStyles(request.Styles, request.HeadingLevels ?? []), OutlineStyle(), NotesConfiguration(request.Notes));
        var automatic = new XElement(Ns.Office + "automatic-styles");
        var masters = new XElement(Ns.Office + "master-styles");
        root.Add(office, automatic, masters);
        var document = new XDocument(root);

        var styles = OdtStyles.Read(document, document);
        var context = new OdtWriteContext(styles, new OdtAutomaticStyles(OdtAutomaticStyles.NamesIn(document)), pictures, new Inventory())
        {
            NotelessDepth = 1,
        };
        var bands = new OdtBandWriter(context, new OdtBlockWriter(context));
        var sections = Effective([.. request.Sections ?? [], request.Page]);
        for (var index = 0; index < sections.Count; index++)
        {
            var layout = LayoutPrefix + (index + 1);
            automatic.Add(OdtTemplatePage.Layout(layout, sections[index]));
            masters.Add(OdtTemplatePage.Master(OdtSectionPlan.MasterNameOf(index, sections.Count), layout, sections[index], bands, context));
        }

        automatic.Add(context.Automatic.Created);
        return document;
    }

    /// <summary>A section that does not declare a band uses the previous one's, kind by kind ("Link to previous").</summary>
    private static List<PageSetupDto> Effective(List<PageSetupDto> sections)
    {
        var resolved = new List<PageSetupDto>();
        foreach (var section in sections)
        {
            resolved.Add(resolved.Count == 0 ? section : section with
            {
                Header = section.Header ?? resolved[^1].Header,
                Footer = section.Footer ?? resolved[^1].Footer,
                FirstHeader = section.FirstHeader ?? resolved[^1].FirstHeader,
                FirstFooter = section.FirstFooter ?? resolved[^1].FirstFooter,
                EvenHeader = section.EvenHeader ?? resolved[^1].EvenHeader,
                EvenFooter = section.EvenFooter ?? resolved[^1].EvenFooter,
            });
        }

        return resolved;
    }

    private static XElement Defaults(StyleSheetDto? sheet)
    {
        var paragraph = OdtFormatWriter.OfStyle(sheet?.Defaults.Paragraph);
        var formatted = OdtFormatWriter.Paragraph(paragraph with { WidowControl = paragraph.WidowControl ?? true });
        var properties = new XElement(
            formatted.Name,
            new XAttribute(Ns.Style + "tab-stop-distance", TabStopDistance),
            new XAttribute(Ns.Style + "writing-mode", "page"),
            formatted.Attributes());
        return new XElement(
            Ns.Style + "default-style",
            new XAttribute(Ns.Style + "family", OdtStyles.ParagraphFamily),
            properties,
            OdtFormatWriter.Text(OdtFormatWriter.OfStyle(sheet?.Defaults.Character), new OdtFonts()));
    }

    /// <param name="headingLevels">A level the sheet has no style for gets the editor's built-in heading.</param>
    private static IEnumerable<XElement> NamedStyles(StyleSheetDto? sheet, List<int> headingLevels)
    {
        var styles = (sheet?.Styles.Values ?? Enumerable.Empty<StyleDefinitionDto>())
            .Where(style => style.Type is "paragraph" or "character").ToList();
        var missing = headingLevels.Select(level => Math.Clamp(level, 1, OdtBlockWriter.MaxHeadingLevel)).Distinct().Order()
            .Where(level => !styles.Any(style => IsHeading(style, level)));
        var written = styles.Concat(missing.Select(HeadingDefinition)).Select(style => Named(style, sheet)).ToList();
        var hasStandard = written.Any(style => style.Attr(Ns.Style + "name") == OdtStyles.DefaultParagraphStyle &&
                                               style.Attr(Ns.Style + "family") == OdtStyles.ParagraphFamily);
        if (!hasStandard) written.Insert(0, Standard());
        return written;
    }

    private static XElement Standard() => new(
        Ns.Style + "style",
        new XAttribute(Ns.Style + "name", OdtStyles.DefaultParagraphStyle),
        new XAttribute(Ns.Style + "family", OdtStyles.ParagraphFamily),
        new XAttribute(Ns.Style + "class", "text"));

    private static bool IsHeading(StyleDefinitionDto style, int level) =>
        style.Type == "paragraph" &&
        (style.Id == OdtWordStyles.HeadingIdOf(level) || string.Equals(style.Name, OdtWordStyles.HeadingNameOf(level), StringComparison.OrdinalIgnoreCase));

    private static StyleDefinitionDto HeadingDefinition(int level)
    {
        var heading = BuiltinStyles.Heading(level);
        return new StyleDefinitionDto(
            heading.Id, heading.Name, "paragraph", heading.QFormat, heading.Hidden, Custom: false,
            BasedOn: heading.BasedOn, Next: heading.Next, UiPriority: heading.UiPriority,
            Paragraph: new StyleParagraphDto(
                SpaceBefore: heading.BeforePt, SpaceAfter: heading.AfterPt, KeepNext: heading.KeepNext ? true : null,
                OutlineLevel: heading.OutlineLevel),
            Character: new StyleCharacterDto(FontSize: OdtFormatWriter.Pt(heading.SizePt), Bold: heading.Bold ? true : null));
    }

    private static XElement Named(StyleDefinitionDto style, StyleSheetDto? sheet)
    {
        var family = style.Type == "paragraph" ? OdtStyles.ParagraphFamily : OdtStyles.TextFamily;
        var name = NameOf(family, style.Id);
        var element = new XElement(Ns.Style + "style", new XAttribute(Ns.Style + "name", name));
        if (OdtWordStyles.NameOf(family, name) is null) element.Add(new XAttribute(Ns.Style + "display-name", style.Name));
        element.Add(new XAttribute(Ns.Style + "family", family));
        if (Related(sheet, style.Type, style.BasedOn) is { } parent) element.Add(new XAttribute(Ns.Style + "parent-style-name", NameOf(family, parent)));
        if (family == OdtStyles.ParagraphFamily && Related(sheet, style.Type, style.Next) is { } next)
        {
            element.Add(new XAttribute(Ns.Style + "next-style-name", NameOf(family, next)));
        }

        if (family == OdtStyles.ParagraphFamily)
        {
            if (style.Paragraph?.OutlineLevel is >= 0 and < WordOutlineLevels and var level)
            {
                element.Add(new XAttribute(Ns.Style + "default-outline-level", level + 1));
            }

            element.Add(OdtFormatWriter.Paragraph(OdtFormatWriter.OfStyle(style.Paragraph)));
        }

        element.Add(OdtFormatWriter.Text(OdtFormatWriter.OfStyle(style.Character), new OdtFonts()));
        foreach (var empty in element.Elements().Where(child => !child.HasAttributes).ToList()) empty.Remove();
        return element;
    }

    private static string NameOf(string family, string id)
    {
        var name = OdtWordStyles.OdfNameOf(family, id);
        return name == id ? OdtStylesPart.SafeName(id) : name;
    }

    /// <summary>The parent or next style, when the sheet has it and of the same type.</summary>
    private static string? Related(StyleSheetDto? sheet, string type, string? id) =>
        id is not null && sheet is not null && sheet.Styles.TryGetValue(id, out var related) && related.Type == type ? id : null;

    /// <summary>Headings without numbers: the default ODF outline numbers them, and the model did not ask for it.</summary>
    private static XElement OutlineStyle() => new(
        Ns.Text + "outline-style",
        new XAttribute(Ns.Style + "name", "Outline"),
        Enumerable.Range(1, OutlineLevels).Select(level => new XElement(
            Ns.Text + "outline-level-style",
            new XAttribute(Ns.Text + "level", level),
            new XAttribute(Ns.Style + "num-format", string.Empty),
            new XElement(
                Ns.Style + "list-level-properties",
                new XAttribute(Ns.Text + "list-level-position-and-space-mode", "label-alignment"),
                new XElement(Ns.Style + "list-level-label-alignment", new XAttribute(Ns.Text + "label-followed-by", "nothing"))))));

    /// <summary>Format, start and restart; LibreOffice counts the start from zero: 0 means "starts at 1".</summary>
    private static IEnumerable<XElement> NotesConfiguration(NotesDto? notes) => new[] { false, true }.Select(endnote =>
    {
        var numbering = OdtNoteCounter.NumberingOf(notes, endnote);
        var format = OdtStylesPart.PageNumberFormats.GetValueOrDefault(OdtNoteCounter.FormatOf(numbering, endnote), "1");
        var restart = numbering?.Restart is { } value && (value != "eachPage" || !endnote) ? NoteRestarts.GetValueOrDefault(value, "document") : "document";
        var element = new XElement(
            Ns.Text + "notes-configuration",
            new XAttribute(Ns.Text + "note-class", endnote ? OdtNoteCounter.Endnote : OdtNoteCounter.Footnote),
            new XAttribute(Ns.Style + "num-format", format),
            new XAttribute(Ns.Text + "start-value", Math.Max(0, (numbering?.Start ?? 1) - 1)));
        if (!endnote) element.Add(new XAttribute(Ns.Text + "footnotes-position", "page"));
        element.Add(new XAttribute(Ns.Text + "start-numbering-at", restart));
        return element;
    });
}

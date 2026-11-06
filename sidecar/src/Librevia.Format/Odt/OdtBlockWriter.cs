using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>The master page a block opens, with the page number it starts at (<c>auto</c> continues).</summary>
internal sealed record OdtPageStart(string MasterPage, string PageNumber);

/// <summary>
/// An editor block → ODF: paragraphs and headings with their direct formatting as an automatic
/// style over the named one, their anchored frames, lists, tables and the page break. The inverse
/// of <see cref="OdtBodyReader"/>: what it writes reads back as the same node.
/// </summary>
internal sealed class OdtBlockWriter
{
    public const int MaxHeadingLevel = 6;

    /// <summary>The reader leaves <c>Standard</c> unsaid, as the DOCX reader leaves <c>Normal</c>.</summary>
    private const string DefaultStyleId = "Normal";

    private readonly OdtWriteContext _context;
    private readonly OdtListWriter _lists;
    private readonly OdtTableWriter _tables;
    private readonly OdtFloatWriter _floats;

    public OdtBlockWriter(OdtWriteContext context)
    {
        _context = context;
        _lists = new OdtListWriter(context, this);
        _tables = new OdtTableWriter(context, this);
        _floats = new OdtFloatWriter(context, this);
    }

    /// <summary>A page break the editor puts before it becomes the block's own <c>fo:break-before</c>.</summary>
    public static bool TakesBreakBefore(Node node) => node.Type is "paragraph" or "heading" or "table";

    /// <summary>A list opens the page in its first paragraph.</summary>
    public static bool OpensPage(Node node) => TakesBreakBefore(node) || node.Type is OdtListLevels.Bullet or OdtListLevels.Ordered;

    /// <summary>The paragraph that only carries a DOCX section's end: ODF puts the section in the next block.</summary>
    public static bool IsSectionMarkOnly(Node node) => Attr.Bool(node, "sectionMark") && (node.Content ?? []).Count == 0;

    /// <param name="flat">Inside a list or a cell, as the reader reads it.</param>
    /// <param name="page">The master page the block opens, which also breaks the page.</param>
    public IEnumerable<XElement> Write(Node node, bool flat = false, bool breakBefore = false, OdtPageStart? page = null)
    {
        switch (node.Type)
        {
            case "paragraph" or "heading" when IsSectionMarkOnly(node):
                return [];
            case "paragraph" or "heading":
                return [Paragraph(node, breakBefore, page)];
            case OdtListLevels.Bullet or OdtListLevels.Ordered:
                return _lists.Write(node, page);
            case "table":
                return [_tables.Write(node, breakBefore, page)];
            case "pageBreak":
                return [PageBreak()];
            case "codeBlock":
                return CodeBlock(node);
            case "horizontalRule":
                return [HorizontalRule()];
            case "tableOfContents":
                _context.NoteLoss(OdtNotices.Indexes);
                return Children(node, flat);
            default:
                _context.NoteLoss(Inventory.Element, node.Type);
                return node.Content?.Any(child => child.Type != "text") == true
                    ? Children(node, flat)
                    : [Paragraph(Node.Of("paragraph", [.. node.Content ?? []]), breakBefore, page)];
        }
    }

    private List<XElement> Children(Node node, bool flat) =>
        [.. (node.Content ?? []).SelectMany(child => Write(child, flat))];

    private XElement Paragraph(Node node, bool breakBefore, OdtPageStart? page)
    {
        var level = node.Type == "heading" ? Math.Clamp(Attr.Int(node, "level") ?? 1, 1, MaxHeadingLevel) : 0;
        var named = _context.NamedStyle(OdtStyles.ParagraphFamily, Attr.String(node, "styleId"))
                    ?? (level > 0 ? _context.HeadingStyle(level) : BodyStyle());

        var ownFont = Attr.String(node, "fontFamily");
        var format = OdtFormatWriter.OfNode(node, ownFont ?? FontOf(named)) with
        {
            BreakBefore = breakBefore && page is null ? ParagraphFormat.Page : null,
        };
        var character = new CharacterFormat(FontFamily: ownFont, FontSizePt: Attr.Points(Attr.String(node, "fontSize")));
        var paragraph = OdtFormatWriter.Paragraph(format);
        if (page is not null) paragraph.SetAttributeValue(Ns.Style + "page-number", page.PageNumber);
        var style = Style(named, paragraph, OdtFormatWriter.Text(character, _context.Styles.Fonts), page);

        var element = new XElement(Ns.Text + (level > 0 ? "h" : "p"));
        if (style is not null) element.Add(new XAttribute(Ns.Text + "style-name", style));
        if (level > 0) element.Add(new XAttribute(Ns.Text + "outline-level", level));
        element.Add(_floats.Write(OdtFloatWriter.FloatsOf(node)));
        element.Add(new OdtInlineWriter(_context, this).Write(node.Content));
        return element;
    }

    /// <summary>An empty paragraph that starts a page: the reader gives the page break and the paragraph.</summary>
    private XElement PageBreak()
    {
        var properties = OdtFormatWriter.Paragraph(ParagraphFormat.Empty with { BreakBefore = ParagraphFormat.Page });
        var style = _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, null, properties);
        return new XElement(Ns.Text + "p", new XAttribute(Ns.Text + "style-name", style));
    }

    /// <summary>One paragraph per line, in the monospaced font and without space between them, with its spaces kept.</summary>
    private IEnumerable<XElement> CodeBlock(Node node)
    {
        var paragraph = OdtFormatWriter.Paragraph(ParagraphFormat.Empty with { SpaceBefore = 0, SpaceAfter = 0 });
        var text = OdtFormatWriter.Text(CharacterFormat.Empty with { FontFamily = OdtInlineWriter.Mono }, _context.Styles.Fonts);
        var style = _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, BodyStyle(), paragraph, text);
        return node.PlainText().Split('\n').Select(line =>
            new XElement(Ns.Text + "p", new XAttribute(Ns.Text + "style-name", style), OdtInlineWriter.PlainText(line)));
    }

    /// <summary>An empty paragraph with a bottom border, as the editor draws the rule.</summary>
    private XElement HorizontalRule()
    {
        var properties = OdtFormatWriter.Paragraph(ParagraphFormat.Empty);
        properties.SetAttributeValue(Ns.Fo + "border-bottom", OdtStyles.DefaultBorder);
        var style = _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, BodyStyle(), properties);
        return new XElement(Ns.Text + "p", new XAttribute(Ns.Text + "style-name", style));
    }

    private string? BodyStyle() => _context.NamedStyle(OdtStyles.ParagraphFamily, DefaultStyleId);

    private string? Style(string? named, XElement paragraph, XElement text, OdtPageStart? page)
    {
        if (page is not null) return _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, named, page.MasterPage, paragraph, text);
        return paragraph.HasAttributes || text.HasAttributes
            ? _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, named, paragraph, text)
            : named;
    }

    /// <summary>The font the named style draws: a line height multiple is measured against it.</summary>
    private string? FontOf(string? named)
    {
        var styles = _context.Styles;
        return styles.Character(styles.Chain(OdtStyles.ParagraphFamily, named), styles.DefaultCharacter()).FontFamily;
    }
}

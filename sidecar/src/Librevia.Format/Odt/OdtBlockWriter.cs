using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// An editor block → ODF: paragraphs and headings with their direct formatting as an automatic
/// style over the named one, lists, tables and the page break. The inverse of
/// <see cref="OdtBodyReader"/>: what it writes reads back as the same node.
/// </summary>
internal sealed class OdtBlockWriter
{
    private const int MaxHeadingLevel = 6;

    /// <summary>The reader leaves <c>Standard</c> unsaid, as the DOCX reader leaves <c>Normal</c>.</summary>
    private const string DefaultStyleId = "Normal";

    private readonly OdtWriteContext _context;
    private readonly OdtListWriter _lists;
    private readonly OdtTableWriter _tables;

    public OdtBlockWriter(OdtWriteContext context)
    {
        _context = context;
        _lists = new OdtListWriter(context, this);
        _tables = new OdtTableWriter(context, this);
    }

    /// <summary>A page break the editor puts before it becomes the block's own <c>fo:break-before</c>.</summary>
    public static bool TakesBreakBefore(Node node) => node.Type is "paragraph" or "heading" or "table";

    /// <param name="flat">Inside a list or a cell, as the reader reads it.</param>
    public IEnumerable<XElement> Write(Node node, bool flat = false, bool breakBefore = false)
    {
        switch (node.Type)
        {
            case "paragraph" or "heading":
                return [Paragraph(node, breakBefore)];
            case OdtListLevels.Bullet or OdtListLevels.Ordered:
                return [_lists.Write(node)];
            case "table":
                return [_tables.Write(node, breakBefore)];
            case "pageBreak":
                return [PageBreak()];
            case "tableOfContents":
                _context.NoteLoss(OdtNotices.Indexes);
                return Children(node, flat);
            default:
                _context.NoteLoss(Inventory.Element, node.Type);
                return node.Content?.Any(child => child.Type != "text") == true
                    ? Children(node, flat)
                    : [Paragraph(Node.Of("paragraph", [.. node.Content ?? []]), breakBefore)];
        }
    }

    private List<XElement> Children(Node node, bool flat) =>
        [.. (node.Content ?? []).SelectMany(child => Write(child, flat))];

    private XElement Paragraph(Node node, bool breakBefore)
    {
        var level = node.Type == "heading" ? Math.Clamp(Attr.Int(node, "level") ?? 1, 1, MaxHeadingLevel) : 0;
        var named = _context.NamedStyle(OdtStyles.ParagraphFamily, Attr.String(node, "styleId"))
                    ?? _context.NamedStyle(OdtStyles.ParagraphFamily, level > 0 ? $"Heading{level}" : DefaultStyleId);

        var ownFont = Attr.String(node, "fontFamily");
        var format = OdtFormatWriter.OfNode(node, ownFont ?? FontOf(named)) with
        {
            BreakBefore = breakBefore ? ParagraphFormat.Page : null,
        };
        var character = new CharacterFormat(FontFamily: ownFont, FontSizePt: Attr.Points(Attr.String(node, "fontSize")));
        var style = Style(named, OdtFormatWriter.Paragraph(format), OdtFormatWriter.Text(character, _context.Styles.Fonts));

        var element = new XElement(Ns.Text + (level > 0 ? "h" : "p"));
        if (style is not null) element.Add(new XAttribute(Ns.Text + "style-name", style));
        if (level > 0) element.Add(new XAttribute(Ns.Text + "outline-level", level));
        element.Add(new OdtInlineWriter(_context).Write(node.Content));
        return element;
    }

    /// <summary>An empty paragraph that starts a page: the reader gives the page break and the paragraph.</summary>
    private XElement PageBreak()
    {
        var properties = OdtFormatWriter.Paragraph(ParagraphFormat.Empty with { BreakBefore = ParagraphFormat.Page });
        var style = _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, null, properties);
        return new XElement(Ns.Text + "p", new XAttribute(Ns.Text + "style-name", style));
    }

    private string? Style(string? named, XElement paragraph, XElement text) =>
        paragraph.HasAttributes || text.HasAttributes
            ? _context.Automatic.Style(OdtAutomaticStyles.Paragraph, OdtStyles.ParagraphFamily, named, paragraph, text)
            : named;

    /// <summary>The font the named style draws: a line height multiple is measured against it.</summary>
    private string? FontOf(string? named)
    {
        var styles = _context.Styles;
        return styles.Character(styles.Chain(OdtStyles.ParagraphFamily, named), styles.DefaultCharacter()).FontFamily;
    }
}

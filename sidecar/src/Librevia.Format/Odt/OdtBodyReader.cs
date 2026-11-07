using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>office:text</c> → the editor's blocks. What has no node here (a section, an index, a shape)
/// keeps its text when it has any, and always goes to the inventory.
/// </summary>
internal sealed class OdtBodyReader
{
    /// <summary>Declarations and settings: no content to show or lose.</summary>
    private static readonly HashSet<string> Silent = new(StringComparer.Ordinal)
    {
        "sequence-decls", "variable-decls", "user-field-decls", "dde-connection-decls", "soft-page-break",
        "calculation-settings", "content-validations", "label-ranges", "named-expressions",
    };

    private readonly OdtReadContext _context;
    private readonly OdtParagraphReader _paragraphs;
    private readonly OdtListReader _lists;
    private readonly OdtTableReader _tables;
    private readonly OdtFrameReader _frames;

    public OdtBodyReader(OdtReadContext context)
    {
        _context = context;
        _paragraphs = new OdtParagraphReader(context);
        _lists = new OdtListReader(context, this);
        _tables = new OdtTableReader(context, this);
        _frames = new OdtFrameReader(context);
    }

    /// <summary>
    /// Each child of <c>office:text</c> that gives nodes is a block, the unit saving keeps byte for
    /// byte or rewrites. A child that gives none stays in the gap before the next block.
    /// </summary>
    public List<OdtSourceBlock> ReadBlocks(XElement text, OdtSectionReader sections)
    {
        var blocks = new List<OdtSourceBlock>();
        var index = 0;
        foreach (var child in text.Elements())
        {
            _context.BeginBlock();
            var opening = sections.OpeningOf(child);
            var nodes = ReadBlock(child, flat: false).ToList();
            var notices = _context.EndBlock();
            if (nodes.Count > 0)
            {
                blocks.Add(new OdtSourceBlock(blocks.Count + 1, index, nodes, notices));
                sections.Add(blocks, blocks[^1], opening);
            }

            index++;
        }

        _context.ModeledSection = null;
        return blocks;
    }

    /// <param name="flat">Inside a list or a cell: see <see cref="OdtParagraphReader"/>.</param>
    public IEnumerable<Node> ReadBlock(XElement element, bool flat)
    {
        var name = element.Name;
        if (name == Ns.Text + "p" || name == Ns.Text + "h") return _paragraphs.Read(element, flat);
        if (name == Ns.Text + "list") return [_lists.Read(element)];
        if (name == Ns.Table + "table") return _tables.Read(element, flat);
        if (name.Namespace == Ns.Draw) return Drawing(element);
        if ((name.Namespace == Ns.Text || name.Namespace == Ns.Table) && Silent.Contains(name.LocalName)) return [];
        return Container(element, flat);
    }

    /// <summary>Sections, indexes and the rest that wraps blocks: their content, and a notice.</summary>
    private IEnumerable<Node> Container(XElement element, bool flat)
    {
        var local = element.Name.LocalName;
        if (element.Name == Ns.Text + "section")
        {
            if (element != _context.ModeledSection) _context.NoteLoss(OdtNotices.Sections);
            return Children(element, flat);
        }

        if (element.Name.Namespace == Ns.Text && OdtNotices.IndexBlocks.Contains(local))
        {
            _context.NoteLoss(OdtNotices.Indexes);
            var body = element.Element(Ns.Text + "index-body");
            return body is null ? [] : Children(body, flat);
        }

        if (element.Name == OdtTrackedChanges.ElementName) return [];

        if (element.Name == Ns.Office + "forms")
        {
            if (element.HasElements) _context.NoteLoss(OdtNotices.Forms);
            return [];
        }

        if (OdtNotices.DroppedInline(element) is { } key)
        {
            _context.NoteLoss(key);
            return [];
        }

        OdtNotices.Element(_context.Inventory, element);
        return Children(element, flat);
    }

    private List<Node> Children(XElement container, bool flat) =>
        [.. container.Elements().SelectMany(child => ReadBlock(child, flat))];

    /// <summary>A frame anchored to the page sits between paragraphs: its image takes one of its own.</summary>
    private IEnumerable<Node> Drawing(XElement element)
    {
        if (element.Name.LocalName != "frame")
        {
            _context.NoteLoss(OdtNotices.ShapeKey(element));
            return [];
        }

        return _frames.Read(element) is { } image ? [Node.Of("paragraph", image)] : [];
    }
}

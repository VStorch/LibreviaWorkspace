using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The new content of <c>office:text</c>: each block the user did not touch goes back as the exact
/// text it had in <c>content.xml</c>; a new or edited one is written by <see cref="OdtBlockWriter"/>.
/// What sits between blocks without being one (declarations, forms, tracked changes, a soft page
/// break) goes back before the block it preceded; the first block's, always first.
/// </summary>
internal sealed class OdtBodyWriter
{
    private const string NoPadding = "0mm";

    private readonly OdtSource _source;
    private readonly OdtWriteContext _context;
    private readonly OdtBlockWriter _blocks;
    private readonly OdtXmlOut _out;
    private readonly string _xml;
    private readonly XmlSpan _text;
    private readonly OdtSectionPlan _sections;
    private readonly Dictionary<string, (OdtSourceBlock Block, int Part)> _byOid = new(StringComparer.Ordinal);
    private readonly HashSet<int> _kept = [];
    private readonly HashSet<int> _placed = [];
    private readonly HashSet<int> _noted = [];
    private readonly HashSet<int> _named = [];
    private readonly Dictionary<int, string> _leading = [];
    private string _prelude = string.Empty;
    private string _trailing = string.Empty;
    private OdtPageStart? _pending;

    public int Preserved { get; private set; }
    public int Rewritten { get; private set; }

    public OdtBodyWriter(OdtSource source, OdtWriteContext context, OdtXmlOut xmlOut, string xml, XmlSpan text, OdtSectionPlan sections)
    {
        _sections = sections;
        _source = source;
        _context = context;
        _blocks = new OdtBlockWriter(context);
        _out = xmlOut;
        _xml = xml;
        _text = text;
        if (text.Children.Count != source.Elements.Count) throw new OdtException("sidecar.error.odtUnreadable");
        foreach (var block in source.Blocks)
        {
            for (var part = 0; part < block.Nodes.Count; part++) _byOid[block.OidAt(part)] = (block, part);
        }
    }

    /// <summary>The text between <c>office:text</c>'s tags.</summary>
    public string Write(Node doc)
    {
        var content = doc.Content ?? [];
        if (_source.Blocks.Count == 0 && content is [{ Type: "paragraph", Content: null } only] && Attr.String(only, "oid") is null)
        {
            return _xml[_text.ContentStart.._text.ContentEnd];
        }

        PlanGaps(content);
        var body = new StringBuilder(_prelude);
        for (var index = 0; index < content.Count;)
        {
            _pending = _sections.Enter(body, index) ?? _pending;
            if (Kept(content, index) is { } block)
            {
                _pending = null;
                Place(body, block);
                body.Append(_xml, Span(block).Start, Span(block).End - Span(block).Start);
                RememberList(block);
                CountNotes(block);
                _kept.Add(block.Number);
                Preserved++;
                index += block.Nodes.Count;
                continue;
            }

            index = Generate(body, content, index);
        }

        _sections.Close(body);
        return body.Append(_trailing).ToString();
    }

    /// <summary>
    /// The gap before each block, and the gaps of deleted blocks carried to the next block still in
    /// the document, so a declaration or a form never goes away with the paragraph after it.
    /// </summary>
    private void PlanGaps(List<Node> content)
    {
        var referenced = content.Select(node => Attr.String(node, "oid")).OfType<string>()
            .Where(_byOid.ContainsKey).Select(oid => _byOid[oid].Block.Number).ToHashSet();
        var carried = new StringBuilder();
        var end = _text.ContentStart;
        foreach (var block in _source.Blocks)
        {
            var gap = _xml[end..Span(block).Start];
            if (block.Number == 1) _prelude = gap;
            else carried.Append(gap);
            end = Span(block).End;
            if (!referenced.Contains(block.Number)) continue;
            _leading[block.Number] = carried.ToString();
            carried.Clear();
        }

        _trailing = carried + _xml[end.._text.ContentEnd];
        if (_source.Blocks.Count > 0) return;
        _prelude = _trailing;
        _trailing = string.Empty;
    }

    private XmlSpan Span(OdtSourceBlock block) => _text.Children[block.ElementIndex];

    /// <summary>The block whose nodes start here, all in order and unchanged, on their first appearance.</summary>
    private OdtSourceBlock? Kept(List<Node> content, int index)
    {
        if (Attr.String(content[index], "oid") is not { } oid || !_byOid.TryGetValue(oid, out var found)) return null;
        var (block, part) = found;
        if (part != 0 || _kept.Contains(block.Number) || index + block.Nodes.Count > content.Count) return null;

        for (var offset = 0; offset < block.Nodes.Count; offset++)
        {
            var node = content[index + offset];
            if (Attr.String(node, "oid") != block.OidAt(offset)) return null;
            if (OdtBlockIdentity.Of(node, _context.Plan.Comments) != OdtBlockIdentity.Of(block.Nodes[offset], _source.Comments ?? [])) return null;
        }

        return block;
    }

    /// <summary>The gap that belongs before the block, once.</summary>
    private void Place(StringBuilder body, OdtSourceBlock block)
    {
        if (!_placed.Add(block.Number)) return;
        if (_leading.TryGetValue(block.Number, out var gap)) body.Append(gap);
    }

    /// <summary>A page break followed by a block that is written too goes into that block's style.</summary>
    private int Generate(StringBuilder body, List<Node> content, int index)
    {
        var node = content[index];
        var next = index + 1 < content.Count ? content[index + 1] : null;
        var merged = node.Type == "pageBreak" && next is not null && OdtBlockWriter.TakesBreakBefore(next) &&
                     Kept(content, index + 1) is null;
        var written = merged ? next! : node;

        _context.SourceFields.Clear();
        foreach (var origin in new[] { node, written }.Distinct())
        {
            if (Origin(origin) is not { } block) continue;
            Place(body, block);
            if (_noted.Add(block.Number)) NoteLosses(block);
            _context.SourceFields.AddRange(_source.Elements[block.ElementIndex].Descendants().Where(OdtNotices.IsTextField));
        }

        if (written.Type == "table" && Origin(written) is { } table && _named.Add(table.Number))
        {
            var element = _source.Elements[table.ElementIndex];
            _context.KeptTableName = element.Attr(Ns.Table + "name");
            _context.KeptCellPadding = CellPaddingOf(element);
        }

        body.Append(_out.Write(_blocks.Write(written, breakBefore: merged, page: TakePending(written))));
        _context.KeptTableName = null;
        _context.KeptCellPadding = null;
        Rewritten += merged ? 2 : 1;
        return index + (merged ? 2 : 1);
    }

    /// <summary>The pending master page goes to the first block that can open one.</summary>
    private OdtPageStart? TakePending(Node node)
    {
        if (_pending is null || OdtBlockWriter.IsSectionMarkOnly(node) || !OdtBlockWriter.OpensPage(node)) return null;
        var page = _pending;
        _pending = null;
        return page;
    }

    /// <summary>The model has no padding: a rewritten table keeps its first cell's, which may be none.</summary>
    private string CellPaddingOf(XElement table)
    {
        var cell = table.Descendants(Ns.Table + "table-cell").FirstOrDefault();
        var properties = _source.Styles.Find("table-cell", cell?.Attr(Ns.Table + "style-name"))?.Properties("table-cell-properties");
        return properties?.Attr(Ns.Fo + "padding") ?? NoPadding;
    }

    private OdtSourceBlock? Origin(Node node) =>
        Attr.String(node, "oid") is { } oid && _byOid.TryGetValue(oid, out var found) ? found.Block : null;

    /// <summary>What the rewritten block held and the model does not: now it is lost.</summary>
    private void NoteLosses(OdtSourceBlock block)
    {
        foreach (var notice in block.Notices.Lost) _context.NoteLoss(notice.Key, notice.Value);
    }

    /// <summary>A kept note still takes its number: the next written one follows it.</summary>
    private void CountNotes(OdtSourceBlock block)
    {
        foreach (var note in block.Nodes.SelectMany(OdtBlockIdentity.Descendants).Where(node => node.Type == "noteRef"))
        {
            _context.Plan.Notes.Next(Attr.String(note, "kind") == OdtNoteCounter.Endnote, _context.Section, Attr.String(note, "mark"));
        }
    }

    /// <summary>A kept list's style, so a written list with the same key continues it.</summary>
    private void RememberList(OdtSourceBlock block)
    {
        foreach (var node in block.Nodes.Where(node => node.Type is OdtListLevels.Bullet or OdtListLevels.Ordered))
        {
            var key = (Attr.Node(node, "numbering") as JsonObject)?["key"]?.GetValue<string>();
            var element = _source.Elements[block.ElementIndex];
            var style = element.Attr(Ns.Text + "style-name") ??
                        _source.Styles.ListStyleOfParagraph(element.Descendants(Ns.Text + "p").FirstOrDefault()?.Attr(Ns.Text + "style-name"));
            if (key is not null && style is not null) _context.ListStyles.TryAdd(key, style);
        }
    }
}

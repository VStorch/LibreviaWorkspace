using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The new content of <c>office:text</c>: each block the user did not touch goes back as the exact
/// text it had in <c>content.xml</c>; a new or edited one is written by <see cref="OdtBlockWriter"/>.
/// What sits between blocks without being one (declarations, forms, tracked changes, a soft page
/// break) goes back before the block it preceded; the first block's, always first. The tracked
/// changes are the exception: their element keeps its place, with only the regions still pointed to
/// and those the rewritten blocks add.
/// </summary>
internal sealed class OdtBodyWriter
{
    private const string NoPadding = "0mm";

    /// <summary>Where the tracked changes go once known: a control character no XML text holds.</summary>
    private const string ChangesPlaceholder = "\u0001";

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
    private readonly XmlSpan? _changesSpan;
    private HashSet<int>? _predicted;
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

        if (source.Changes.Element is { } changes) _changesSpan = text.Children[source.Elements.IndexOf(changes)];
    }

    private OdtChangeWriter Changes => _context.Changes!;

    /// <summary>The text between <c>office:text</c>'s tags.</summary>
    public string Write(Node doc)
    {
        var content = doc.Content ?? [];
        if (_source.Blocks.Count == 0 && content is [{ Type: "paragraph", Content: null } only] && Attr.String(only, "oid") is null &&
            !Changes.ChangesFlag)
        {
            return _xml[_text.ContentStart.._text.ContentEnd];
        }

        PlanGaps(content);
        Predict(content);
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
        var written = body.Append(_trailing).ToString();
        var changes = Changes.Write(_xml, _changesSpan, Referenced(_kept), _out);
        return written.Contains(ChangesPlaceholder, StringComparison.Ordinal)
            ? written.Replace(ChangesPlaceholder, changes, StringComparison.Ordinal)
            : changes + written;
    }

    /// <summary>
    /// The blocks kept as they are, decided before writing, so a rewritten block knows which region
    /// ids are taken. The blocks from a range's start to its end stay or go together, those between
    /// without a mark of their own too: half a range would not close, or would cover rewritten text.
    /// </summary>
    private void Predict(List<Node> content)
    {
        var kept = new HashSet<int>();
        for (var index = 0; index < content.Count;)
        {
            if (Kept(content, index) is { } block && kept.Add(block.Number)) index += block.Nodes.Count;
            else index++;
        }

        var linked = _source.Blocks
            .SelectMany(block => OdtTrackedChanges.ReferencesIn(_source.Elements[block.ElementIndex]).Select(id => (id, block.Number)))
            .GroupBy(link => link.id, link => link.Number)
            .Select(group => Enumerable.Range(group.Min(), group.Max() - group.Min() + 1).ToList())
            .Where(group => group.Count > 1).ToList();
        for (var split = true; split;)
        {
            split = false;
            foreach (var group in linked.Where(group => group.Any(kept.Contains) && !group.All(kept.Contains)))
            {
                kept.ExceptWith(group);
                split = true;
            }
        }

        _predicted = kept;
        Changes.Reserve(Referenced(kept));
    }

    /// <summary>The regions what goes back as it is points to: the kept blocks and what sits between blocks.</summary>
    private HashSet<string> Referenced(IReadOnlySet<int> kept)
    {
        var blockIndexes = _source.Blocks.Select(block => block.ElementIndex).ToHashSet();
        var raw = _source.Blocks.Where(block => kept.Contains(block.Number)).Select(block => _source.Elements[block.ElementIndex])
            .Concat(_source.Elements.Where((element, index) => !blockIndexes.Contains(index) && element != _source.Changes.Element));
        return raw.SelectMany(OdtTrackedChanges.ReferencesIn).ToHashSet(StringComparer.Ordinal);
    }

    /// <summary>The file's text between two offsets, with the tracked changes' element left for later.</summary>
    private string Gap(int start, int end)
    {
        if (_changesSpan is not { } changes || changes.Start < start || changes.End > end) return _xml[start..end];
        return _xml[start..changes.Start] + ChangesPlaceholder + _xml[changes.End..end];
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
            var gap = Gap(end, Span(block).Start);
            if (block.Number == 1) _prelude = gap;
            else carried.Append(gap);
            end = Span(block).End;
            if (!referenced.Contains(block.Number)) continue;
            _leading[block.Number] = carried.ToString();
            carried.Clear();
        }

        _trailing = carried + Gap(end, _text.ContentEnd);
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
        if (_predicted?.Contains(block.Number) == false) return null;
        if (index > 0 && OdtBlockWriter.JoinsNext(content[index - 1]) && OdtBlockWriter.IsParagraph(content[index])) return null;

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

    /// <summary>
    /// A page break followed by a block that is written too goes into that block's style; paragraphs
    /// joined by deleted paragraph marks are written as one.
    /// </summary>
    private int Generate(StringBuilder body, List<Node> content, int index)
    {
        var node = content[index];
        var next = index + 1 < content.Count ? content[index + 1] : null;
        var merged = node.Type == "pageBreak" && next is not null && OdtBlockWriter.TakesBreakBefore(next) &&
                     Kept(content, index + 1) is null;
        var start = merged ? index + 1 : index;
        var chain = content.GetRange(start, OdtBlockWriter.ChainLength(content, start));
        var written = chain[0];

        _context.SourceFields.Clear();
        foreach (var origin in chain.Prepend(node).Distinct())
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

        var elements = _blocks.Write(chain, breakBefore: merged, page: TakePending(written)).ToList();
        var end = start + chain.Count;
        if (!ContinuesText(content, end)) Changes.Abandon();
        body.Append(_out.Write(elements));
        _context.KeptTableName = null;
        _context.KeptCellPadding = null;
        Rewritten += end - index;
        return end;
    }

    /// <summary>Whether the next block is written too, with a paragraph first, so an open range can close in it.</summary>
    private bool ContinuesText(List<Node> content, int index) =>
        index < content.Count && content[index].Type is "paragraph" or "heading" or OdtListLevels.Bullet or OdtListLevels.Ordered &&
        Kept(content, index) is null;

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

using System.Text;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>An insertion range being written: its region, and the <c>text:change-start</c> that opened it.</summary>
/// <param name="Paragraph">The paragraph whose mark it leaves open, once written.</param>
internal sealed record OdtOpenChange(Mark Mark, string Id, XElement Start, XElement? Paragraph = null);

/// <summary>
/// The regions of a save. A rewritten block takes the region id its revision was read with, when
/// no kept block points to it and no other range took it; otherwise one the file does not use. An
/// inserted paragraph mark leaves its range open (<see cref="Pending"/>) for the next paragraph
/// written to close: a range that nobody continues closes where it opened.
/// </summary>
internal sealed class OdtChangeWriter
{
    private const string IdPrefix = "LvChange";

    private readonly OdtTrackedChanges _source;
    private readonly bool? _wanted;
    private readonly HashSet<string> _taken = new(StringComparer.Ordinal);
    private readonly HashSet<string> _reserved = new(StringComparer.Ordinal);
    private readonly List<XElement> _written = [];
    private int _next = 1;

    /// <param name="wanted">The recording flag the model asks for; null leaves the file's.</param>
    public OdtChangeWriter(OdtTrackedChanges source, XElement root, bool? wanted)
    {
        _source = source;
        _wanted = wanted;
        foreach (var element in root.Descendants())
        {
            if (element.Attr(Ns.Xml + "id") is { } xmlId) _taken.Add(xmlId);
            if (element.Name == Ns.Text + "changed-region" && element.Attr(Ns.Text + "id") is { } id) _taken.Add(id);
        }
    }

    public OdtOpenChange? Pending { get; private set; }

    /// <summary>Whether the model turns recording on or off.</summary>
    public bool ChangesFlag => _wanted is { } wanted && wanted != _source.Tracking;

    /// <summary>The ids kept blocks point to: a rewritten one never takes them.</summary>
    public void Reserve(IEnumerable<string> ids) => _reserved.UnionWith(ids);

    public OdtOpenChange Open(Mark insertion)
    {
        var id = Region(insertion, OdtChangedRegion.Insertion, []);
        return new OdtOpenChange(insertion, id, OdtTrackedChanges.Reference("change-start", id));
    }

    public static XElement Close(OdtOpenChange change) => OdtTrackedChanges.Reference("change-end", change.Id);

    public XElement Point(Mark deletion, IEnumerable<XElement> paragraphs) =>
        OdtTrackedChanges.Reference("change", Region(deletion, OdtChangedRegion.Deletion, paragraphs));

    public void Carry(OdtOpenChange change) => Pending = change;

    /// <summary>The paragraph just written, when it left a range open.</summary>
    public void Anchor(XElement paragraph)
    {
        if (Pending is { Paragraph: null } pending) Pending = pending with { Paragraph = paragraph };
    }

    public OdtOpenChange? TakePending()
    {
        var pending = Pending;
        Pending = null;
        return pending;
    }

    /// <summary>The paragraph mark stays as it is: the range closes at the end of its paragraph's text.</summary>
    public void Abandon()
    {
        if (TakePending() is not { } pending) return;
        if (pending.Paragraph is { } paragraph) paragraph.Add(Close(pending));
        else pending.Start.AddAfterSelf(Close(pending));
    }

    /// <summary>Before a text of its own (a cell, a note): nothing open crosses into it.</summary>
    public OdtOpenChange? Suspend() => TakePending();

    public void Resume(OdtOpenChange? outer)
    {
        Abandon();
        Pending = outer;
    }

    private string Region(Mark mark, string kind, IEnumerable<XElement> content)
    {
        var id = IdOf(Attr.MarkString(mark, "rid"), kind);
        _written.Add(new XElement(
            Ns.Text + "changed-region",
            new XAttribute(Ns.Xml + "id", id),
            new XAttribute(Ns.Text + "id", id),
            new XElement(Ns.Text + kind, ChangeInfo(mark), content)));
        return id;
    }

    private string IdOf(string? rid, string kind)
    {
        if (rid is not null && _source.Find(rid) is { } read && read.Kind == kind && read.Id == rid &&
            _reserved.Add(rid))
        {
            return rid;
        }

        string id;
        while (!_taken.Add(id = $"{IdPrefix}{_next}")) _next++;
        return id;
    }

    private static XElement ChangeInfo(Mark mark)
    {
        var info = new XElement(Ns.Office + "change-info");
        if (Attr.MarkString(mark, "author") is { } author) info.Add(new XElement(Ns.Dc + "creator", author));
        if (Attr.MarkString(mark, "date") is { } date) info.Add(new XElement(Ns.Dc + "date", date));
        return info;
    }

    /// <summary>
    /// The new <c>text:tracked-changes</c>: the regions <paramref name="referenced"/> as the file has
    /// them, then the written ones. Unchanged, it is the file's text; empty and not recording, nothing.
    /// </summary>
    /// <param name="span">The original element's, or null when the file has none.</param>
    public string Write(string xml, XmlSpan? span, IReadOnlySet<string> referenced, OdtXmlOut xmlOut)
    {
        var element = _source.Element;
        var current = _source.Tracking;
        var wanted = _wanted ?? current;
        var kept = _source.Regions.Where(region => referenced.Contains(region.Id)).ToList();
        if (span is not null && kept.Count == _source.Regions.Count && _written.Count == 0 && wanted == current)
        {
            return xml[span.Start..span.End];
        }

        if (kept.Count + _written.Count == 0 && !wanted) return string.Empty;

        var shell = new XElement(OdtTrackedChanges.ElementName, element?.Attributes() ?? []);
        if (wanted != current || element is null)
        {
            shell.SetAttributeValue(OdtTrackedChanges.TrackAttribute, wanted ? null : "false");
        }

        var text = new StringBuilder(xmlOut.Open(shell));
        if (span is not null) AppendKept(text, xml, span, kept);
        text.Append(xmlOut.Write(_written));
        return text.Append(xmlOut.Close(shell)).ToString();
    }

    /// <summary>The spans stop at the blocks: the element's own text gives its regions'.</summary>
    private void AppendKept(StringBuilder text, string xml, XmlSpan span, List<OdtChangedRegion> kept)
    {
        var own = xml[span.Start..span.End];
        var children = OdtXmlSpans.Root(own).Children;
        var elements = _source.Element!.Elements().ToList();
        foreach (var region in kept)
        {
            var child = children[elements.IndexOf(region.Element)];
            text.Append(own, child.Start, child.End - child.Start);
        }
    }
}

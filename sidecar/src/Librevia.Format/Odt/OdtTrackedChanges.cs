using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A <c>text:changed-region</c>: its kind (<c>insertion</c>, <c>deletion</c>, <c>format-change</c>)
/// and the mark it gives, with the author and date as the file wrote them and the region's id as
/// <c>rid</c>, as <see cref="Revisions"/> gives a DOCX revision.
/// </summary>
internal sealed record OdtChangedRegion(string Id, string Kind, Mark Mark, XElement Element)
{
    public const string Insertion = "insertion";
    public const string Deletion = "deletion";

    /// <summary>What a deletion took away: the blocks after its <c>office:change-info</c>.</summary>
    public IEnumerable<XElement> Deleted =>
        Element.Element(Ns.Text + Kind)?.Elements().Where(child => child.Name != Ns.Office + "change-info") ?? [];
}

/// <summary>
/// The <c>text:tracked-changes</c> of <c>office:text</c>. Its <c>text:track-changes</c> is the
/// recording flag; absent means on, as ODF and LibreOffice read it.
/// </summary>
internal sealed class OdtTrackedChanges
{
    public static readonly XName ElementName = Ns.Text + "tracked-changes";
    public static readonly XName TrackAttribute = Ns.Text + "track-changes";
    private static readonly XName ChangeId = Ns.Text + "change-id";

    private static readonly string[] MarkAttributes = ["author", "date", "rid"];

    private static readonly HashSet<string> ChangeMarks = new(StringComparer.Ordinal) { "change", "change-start", "change-end" };

    private readonly Dictionary<string, OdtChangedRegion> _byId = new(StringComparer.Ordinal);

    public static readonly OdtTrackedChanges None = new(null);

    private OdtTrackedChanges(XElement? element)
    {
        Element = element;
        foreach (var region in element?.Elements(Ns.Text + "changed-region") ?? [])
        {
            var id = region.Attr(Ns.Text + "id") ?? region.Attr(Ns.Xml + "id");
            var kind = region.Elements().FirstOrDefault(child => child.Name.Namespace == Ns.Text)?.Name.LocalName;
            if (id is null || kind is null) continue;
            var found = new OdtChangedRegion(id, kind, MarkOf(id, kind, region.Element(Ns.Text + kind)), region);
            Regions.Add(found);
            _byId.TryAdd(id, found);
            if (region.Attr(Ns.Xml + "id") is { } xmlId) _byId.TryAdd(xmlId, found);
        }
    }

    public XElement? Element { get; }

    public List<OdtChangedRegion> Regions { get; } = [];

    public bool Tracking => Element is not null && Element.Attr(TrackAttribute) != "false";

    public static OdtTrackedChanges Read(XDocument content) =>
        new(content.Root?.Element(Ns.Office + "body")?.Element(Ns.Office + "text")?.Element(ElementName));

    public OdtChangedRegion? Find(string? id) => id is null ? null : _byId.GetValueOrDefault(id);

    public static bool IsChangeMark(XElement element) =>
        element.Name.Namespace == Ns.Text && ChangeMarks.Contains(element.Name.LocalName);

    /// <summary>The regions an element points to, its descendants included.</summary>
    public static IEnumerable<string> ReferencesIn(XElement element) =>
        element.DescendantsAndSelf().Where(IsChangeMark).Select(mark => mark.Attr(ChangeId)).OfType<string>();

    public static XElement Reference(string localName, string id) =>
        new(Ns.Text + localName, new XAttribute(ChangeId, id));

    private static Mark MarkOf(string id, string kind, XElement? change)
    {
        var info = change?.Element(Ns.Office + "change-info");
        var attrs = new Dictionary<string, JsonNode?>();
        if (info?.Element(Ns.Dc + "creator")?.Value is { } author) attrs["author"] = author;
        if (info?.Element(Ns.Dc + "date")?.Value is { } date) attrs["date"] = date;
        attrs["rid"] = id;
        return new Mark { Type = kind == OdtChangedRegion.Deletion ? Revisions.Deletion : Revisions.Insertion, Attrs = attrs };
    }

    /// <summary>The paragraph mark's revision (<c>markRevision</c>) of a text revision.</summary>
    public static JsonObject BlockRevisionOf(Mark mark)
    {
        var result = new JsonObject { ["kind"] = mark.Type == Revisions.Deletion ? "del" : "ins" };
        foreach (var name in MarkAttributes)
        {
            if (Attr.MarkString(mark, name) is { } value) result[name] = value;
        }

        return result;
    }

    /// <summary>The inverse of <see cref="BlockRevisionOf"/>, or null without a known kind.</summary>
    public static Mark? MarkOfBlock(JsonNode? revision)
    {
        if (revision is not JsonObject block || block["kind"]?.GetValue<string>() is not { } kind || kind is not ("ins" or "del")) return null;
        var attrs = new Dictionary<string, JsonNode?>();
        foreach (var name in MarkAttributes)
        {
            if (block[name]?.GetValue<string>() is { } value) attrs[name] = value;
        }

        return new Mark { Type = kind == "del" ? Revisions.Deletion : Revisions.Insertion, Attrs = attrs };
    }

    /// <summary>One revision: same kind, author, date and id, as Word merges them.</summary>
    public static bool Same(Mark a, Mark b) =>
        a.Type == b.Type && MarkAttributes.All(name => Attr.MarkString(a, name) == Attr.MarkString(b, name));
}

using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>A section as read: the block it starts at, the master page it uses and its setup.</summary>
/// <param name="Encoded">Whether the start block's own XML says it: a master page or columns.</param>
/// <param name="Container">The <c>text:section</c> of a section with columns, reopened as it was when they stay.</param>
internal sealed record OdtReadSection(int StartBlock, string? Master, PageSetupDto Setup, bool Encoded = false, XElement? Container = null);

/// <summary>
/// The body's sections, the inverse of <see cref="OdtSectionPlan"/>: a top-level block whose style
/// names a master page starts a section on the next page; a top-level <c>text:section</c> with
/// columns is a continuous section, and the blocks after it another. As in a DOCX, the block that
/// closes a section carries its <c>sectionBreak</c> (<c>s1</c>, <c>s2</c>…), on its last paragraph.
/// A master page further inside a block, or a section the model cannot hold, keeps the notice.
/// </summary>
internal sealed class OdtSectionReader(OdtReadContext context, Inventory notices)
{
    public const string NextPage = "nextPage";
    public const string Continuous = "continuous";
    private const string SectionIdPrefix = "s";

    private sealed record Start(int Block, string? Master, string Kind, int? PageNumber, ColumnsDto? Columns, bool Encoded, XElement? Container);

    /// <summary>What the block about to be read starts, decided from its element.</summary>
    public sealed record Opening(string? Master, int? PageNumber, ColumnsDto? Columns, XElement? Container = null);

    private readonly List<Start> _starts = [];
    private readonly Dictionary<string, PageSetupDto> _pages = new(StringComparer.Ordinal);
    private string? _master;
    private bool _afterColumns;

    /// <summary>Called inside the block, so a notice goes to it.</summary>
    public Opening OpeningOf(XElement element)
    {
        var starter = StarterOf(element);
        var style = starter is null ? null : StyleOf(starter);
        var master = style?.MasterPage is { Length: > 0 } named ? named : null;
        var pageNumber = master is null ? null : PageNumberOf(style!, starter!);
        foreach (var inner in element.DescendantsAndSelf().Where(inner => inner != starter && IsStarter(inner)))
        {
            if (StyleOf(inner)?.MasterPage is { Length: > 0 }) context.NoteLoss(OdtNotices.Sections);
        }

        var columns = element.Name == Ns.Text + "section" ? ColumnsOf(element) : null;
        var modeled = columns is not null && !element.Descendants(Ns.Text + "section").Any() &&
                      !element.Descendants().Any(inner => IsStarter(inner) && StyleOf(inner)?.MasterPage is { Length: > 0 });
        context.ModeledSection = modeled ? element : null;
        return modeled ? new Opening(master, pageNumber, columns, element) : new Opening(master, pageNumber, null);
    }

    /// <summary>After the block is read: a new section closes the previous one at the block before.</summary>
    public void Add(List<OdtSourceBlock> blocks, OdtSourceBlock block, Opening opening)
    {
        var kind = _starts.Count == 0 ? NextPage
            : opening.Master is not null ? NextPage
            : opening.Columns is not null || _afterColumns ? Continuous
            : null;
        _afterColumns = opening.Columns is not null;
        if (kind is null) return;

        var first = _starts.Count == 0;
        if (!first) MarkBreak(blocks[^2], SectionIdPrefix + _starts.Count);
        _master = opening.Master ?? _master;
        var pageNumber = first && opening.PageNumber == 1 ? null : opening.PageNumber;
        _starts.Add(new Start(block.Number, _master, kind, pageNumber, opening.Columns, opening.Master is not null || opening.Columns is not null, opening.Container));
    }

    /// <summary>The last is the page; one section alone reads as it always did, without a start.</summary>
    public List<OdtReadSection> Sections()
    {
        if (_starts.Count == 0) return [new OdtReadSection(0, null, PageOf(null))];
        var single = _starts.Count == 1;
        return [.. _starts.Select((start, index) => new OdtReadSection(start.Block, start.Master, PageOf(start.Master) with
        {
            Id = index < _starts.Count - 1 ? SectionIdPrefix + (index + 1) : null,
            Start = single ? null : start.Kind,
            PageNumberStart = PageReader.StartElement(start.PageNumber),
            Columns = start.Columns,
        }, start.Encoded, start.Container))];
    }

    private PageSetupDto PageOf(string? master)
    {
        var key = master ?? string.Empty;
        if (!_pages.TryGetValue(key, out var page)) _pages[key] = page = OdtPageReader.Read(context.Package, master, notices);
        return page;
    }

    /// <summary>On the closing block's last paragraph, where the editor looks for it.</summary>
    private static void MarkBreak(OdtSourceBlock block, string id)
    {
        var last = block.Nodes[^1];
        var target = last.Type is "paragraph" or "heading"
            ? last
            : OdtBlockIdentity.Descendants(last).LastOrDefault(node => node.Type is "paragraph" or "heading") ?? last;
        target.With("sectionBreak", id);
    }

    /// <summary>The element whose style opens the block's page: a list's first paragraph.</summary>
    private static XElement? StarterOf(XElement element)
    {
        if (IsStarter(element)) return element;
        return element.Name == Ns.Text + "list"
            ? element.Descendants().FirstOrDefault(inner => inner.Name == Ns.Text + "p" || inner.Name == Ns.Text + "h")
            : null;
    }

    private static bool IsStarter(XElement element) =>
        element.Name == Ns.Text + "p" || element.Name == Ns.Text + "h" || element.Name == Ns.Table + "table";

    private OdtStyle? StyleOf(XElement element) => element.Name == Ns.Table + "table"
        ? context.Styles.Find("table", element.Attr(Ns.Table + "style-name"))
        : context.Styles.Find(OdtStyles.ParagraphFamily, element.Attr(Ns.Text + "style-name"));

    private static int? PageNumberOf(OdtStyle style, XElement starter)
    {
        var kind = starter.Name == Ns.Table + "table" ? "table-properties" : "paragraph-properties";
        return OdtLength.Integer(style.Properties(kind)?.Attr(Ns.Style + "page-number"));
    }

    /// <summary>More than one column, or <c>null</c>.</summary>
    private ColumnsDto? ColumnsOf(XElement section)
    {
        var columns = context.Styles.Find(OdtAutomaticStyles.SectionFamily, section.Attr(Ns.Text + "style-name"))?
            .Properties("section-properties")?.Element(Ns.Style + "columns");
        if (columns is null || OdtLength.Integer(columns.Attr(Ns.Fo + "column-count")) is not ({ } count and > 1)) return null;
        var separator = columns.Element(Ns.Style + "column-sep");
        return new ColumnsDto(
            count,
            OdtLength.RoundedMillimetres(columns.Attr(Ns.Fo + "column-gap")) ?? 0,
            separator is not null && separator.Attr(Ns.Style + "style") != "none");
    }
}

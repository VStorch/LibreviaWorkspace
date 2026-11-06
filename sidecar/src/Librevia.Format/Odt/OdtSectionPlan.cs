using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The model's sections over the body: each closes at the block that carries its
/// <c>sectionBreak</c>, and the last is the page's. The block that opens a section opens its
/// master page (<c>Section1</c>, <c>Section2</c>… and, for the page's, the last one, <c>Standard</c>,
/// whose layout <see cref="OdtStylesPart"/> keeps up with the page, as <see cref="OdtTemplate"/>
/// names them), unless
/// it continues on the same sheet; columns go in a <c>text:section</c>, which can start mid-sheet.
/// A section whose master page the package lacks keeps the previous page, and is declared.
/// </summary>
internal sealed class OdtSectionPlan
{
    public const string PageMaster = "Standard";

    private const string Continuous = "continuous";
    private const string AutoPageNumber = "auto";
    private const double DefaultColumnGapMm = 12.7;

    private readonly int[] _sectionOf;
    private readonly List<PageSetupDto> _sections;
    private readonly bool _numbered;
    private readonly ISet<string> _masters;
    private readonly string _pageMaster;
    private readonly OdtWriteContext _context;
    private readonly OdtXmlOut _out;
    private int _current = -1;
    private XElement? _open;

    /// <param name="pageMaster">The master page of the page setup, the first block's in an opened file.</param>
    public OdtSectionPlan(DocumentModelDto model, ISet<string> masters, string? pageMaster, OdtWriteContext context, OdtXmlOut xmlOut)
    {
        _pageMaster = pageMaster ?? PageMaster;
        _sections = [.. model.Sections ?? [], model.Page];
        _numbered = model.Sections is { Count: > 0 } || model.Page.PageNumberStart.ValueKind == JsonValueKind.Number;
        _masters = masters;
        _context = context;
        _out = xmlOut;
        _sectionOf = SectionsOf(model.Doc.Content ?? [], model.Sections ?? []);
    }

    public static string MasterNameOf(int index, int count) => index == count - 1 ? PageMaster : $"Section{index + 1}";

    /// <summary>The master page the block opens, or <c>null</c> when it stays in the section it is in.</summary>
    public OdtPageStart? Enter(StringBuilder body, int index)
    {
        if (index >= _sectionOf.Length || _sectionOf[index] == _current) return null;
        var previous = _current;
        _current = _sectionOf[index];
        _context.Section = _current;
        Close(body);

        var section = _sections[_current];
        if (section.Columns is { Count: > 1 } columns) Open(body, columns);
        if (!_numbered || (previous >= 0 && section.Start == Continuous)) return null;
        var master = _current == _sections.Count - 1 ? _pageMaster : MasterNameOf(_current, _sections.Count);
        if (!_masters.Contains(master))
        {
            _context.NoteLoss(OdtNotices.Sections);
            return null;
        }

        return new OdtPageStart(master, PageNumberOf(section, first: previous < 0));
    }

    public void Close(StringBuilder body)
    {
        if (_open is null) return;
        body.Append(_out.Close(_open));
        _open = null;
    }

    private void Open(StringBuilder body, ColumnsDto columns)
    {
        var layout = new XElement(
            Ns.Style + "columns",
            new XAttribute(Ns.Fo + "column-count", columns.Count),
            new XAttribute(Ns.Fo + "column-gap", OdtFormatWriter.Mm(columns.SpaceMm > 0 ? columns.SpaceMm : DefaultColumnGapMm)!));
        if (columns.Separator)
        {
            layout.Add(new XElement(
                Ns.Style + "column-sep",
                new XAttribute(Ns.Style + "width", "0.2mm"),
                new XAttribute(Ns.Style + "color", "#000000"),
                new XAttribute(Ns.Style + "height", "100%"),
                new XAttribute(Ns.Style + "vertical-align", "top")));
        }

        var properties = new XElement(Ns.Style + "section-properties", new XAttribute(Ns.Text + "dont-balance-text-columns", "false"), layout);
        var style = _context.Automatic.Style(OdtAutomaticStyles.Section, OdtAutomaticStyles.SectionFamily, null, properties);
        _open = new XElement(Ns.Text + "section", new XAttribute(Ns.Text + "style-name", style), new XAttribute(Ns.Text + "name", _context.NewSectionName()));
        body.Append(_out.Open(_open));
    }

    private static string PageNumberOf(PageSetupDto section, bool first) =>
        PageReader.TryStartOf(section, out var start) && start is { } value
            ? value.ToString(CultureInfo.InvariantCulture)
            : first ? "1" : AutoPageNumber;

    /// <summary>From the end: each block belongs to the section of the next break at or after it.</summary>
    private static int[] SectionsOf(List<Node> blocks, List<PageSetupDto> sections)
    {
        var indexById = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var index = 0; index < sections.Count; index++)
        {
            if (sections[index].Id is { } id) indexById.TryAdd(id, index);
        }

        var result = new int[blocks.Count];
        var current = sections.Count;
        for (var index = blocks.Count - 1; index >= 0; index--)
        {
            if (BreakIn(blocks[index]) is { } id && indexById.TryGetValue(id, out var known)) current = known;
            result[index] = current;
        }

        return result;
    }

    /// <summary>Its own, or that of a paragraph inside it: in a list the mark is on the inner paragraph.</summary>
    private static string? BreakIn(Node node)
    {
        if (Attr.String(node, "sectionBreak") is { Length: > 0 } own) return own;
        if (node.Type is "paragraph" or "heading") return null;
        return (node.Content ?? []).Select(BreakIn).FirstOrDefault(found => found is not null);
    }
}

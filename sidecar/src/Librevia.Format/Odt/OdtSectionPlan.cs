using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The model's sections over the body: each closes at the block that carries its
/// <c>sectionBreak</c>, and the last is the page's. The block that opens a section opens its
/// master page (<see cref="OdtSectionMasters"/>; in a new file <c>Section1</c>, <c>Section2</c>… and
/// <c>Standard</c> for the page, as <see cref="OdtTemplate"/> names them), unless it continues on
/// the same sheet; columns go in a <c>text:section</c>, which can start mid-sheet. A section whose
/// master page the package lacks, or that changes it without a new sheet, is declared.
/// </summary>
internal sealed class OdtSectionPlan
{
    public const string PageMaster = "Standard";

    private const string Continuous = "continuous";
    private const string AutoPageNumber = "auto";
    private const double DefaultColumnGapMm = 12.7;
    private const double ColumnGapToleranceMm = 0.05;

    private readonly int[] _sectionOf;
    private readonly List<PageSetupDto> _sections;
    private readonly bool _numbered;
    private readonly OdtSectionMasters _masters;
    private readonly OdtWriteContext _context;
    private readonly OdtXmlOut _out;
    private int _current = -1;
    private XElement? _open;

    public OdtSectionPlan(DocumentModelDto model, OdtSectionMasters masters, OdtWriteContext context, OdtXmlOut xmlOut)
    {
        _sections = [.. model.Sections ?? [], model.Page];
        _numbered = model.Sections is { Count: > 0 } || model.Page.PageNumberStart.ValueKind == JsonValueKind.Number;
        _masters = masters;
        _context = context;
        _out = xmlOut;
        _sectionOf = SectionsOf(model.Doc.Content ?? [], model.Sections ?? []);
    }

    public static string MasterNameOf(int index, int count) => index == count - 1 ? PageMaster : $"Section{index + 1}";

    /// <summary>The master page the block opens, or <c>null</c> when it stays in the section it is in.</summary>
    /// <param name="raw">A block kept as it was: its XML already holds its master page or columns.</param>
    public OdtPageStart? Enter(StringBuilder body, int index, bool raw = false)
    {
        if (index >= _sectionOf.Length || _sectionOf[index] == _current) return null;
        var previous = _current;
        _current = _sectionOf[index];
        _context.Section = _current;
        Close(body);

        var section = _sections[_current];
        if (!raw && section.Columns is { Count: > 1 } columns) Open(body, columns, KeptContainer(_current, columns));
        if (!OpensMaster(_current, first: previous < 0))
        {
            if (previous >= 0 && !OdtStylesPart.SameLayout(_sections[previous], section)) _context.NoteLoss(OdtNotices.Sections);
            return null;
        }

        var master = _masters.Names[_current];
        if (!_masters.Available.Contains(master))
        {
            _context.NoteLoss(OdtNotices.Sections);
            return null;
        }

        return new OdtPageStart(master, PageNumberOf(section, first: previous < 0));
    }

    private bool OpensMaster(int section, bool first) =>
        first ? _numbered || _masters.Names[section] != OdtPageReader.DefaultMaster : _sections[section].Start != Continuous;

    /// <summary>The model's section that starts at this block, or <c>null</c>.</summary>
    public int? StartAt(int index) =>
        index < _sectionOf.Length && (index == 0 || _sectionOf[index] != _sectionOf[index - 1]) ? _sectionOf[index] : null;

    /// <summary>
    /// Whether a block kept as it was says what the plan wants here: the read section that started at
    /// it (<paramref name="read"/>) is the one starting here, on the same master page and opening the
    /// same way; or neither starts, or the one that does needs nothing in the block's XML.
    /// </summary>
    public bool Keeps(int index, int? read)
    {
        var section = StartAt(index);
        if (section is null) return read is null || !_masters.Read[read.Value].Encoded;
        if (read is null) return !Writes(section.Value, first: index == 0);

        var known = _masters.Read[read.Value];
        return _masters.ReadOf[section.Value] == read && _masters.Names[section.Value] == (known.Master ?? OdtPageReader.DefaultMaster) &&
               SameOpening(known.Setup, _sections[section.Value]);
    }

    private bool Writes(int section, bool first) => _sections[section].Columns is { Count: > 1 } || OpensMaster(section, first);

    private static bool SameOpening(PageSetupDto known, PageSetupDto wanted)
    {
        PageReader.TryStartOf(known, out var knownStart);
        PageReader.TryStartOf(wanted, out var wantedStart);
        return (known.Start ?? OdtSectionReader.NextPage) == (wanted.Start ?? OdtSectionReader.NextPage) && knownStart == wantedStart &&
               SameColumns(known.Columns, wanted.Columns);
    }

    private static bool SameColumns(ColumnsDto? known, ColumnsDto? wanted)
    {
        if (known is not { Count: > 1 } || wanted is not { Count: > 1 }) return known is not { Count: > 1 } && wanted is not { Count: > 1 };
        return known.Count == wanted.Count && known.Separator == wanted.Separator && Math.Abs(known.SpaceMm - wanted.SpaceMm) < ColumnGapToleranceMm;
    }

    public void Close(StringBuilder body)
    {
        if (_open is null) return;
        body.Append(_out.Close(_open));
        _open = null;
    }

    /// <summary>The file's own <c>text:section</c> start tag, when the section's columns did not change.</summary>
    private XElement? KeptContainer(int section, ColumnsDto columns) =>
        _masters.ReadOf[section] is { } read && _masters.Read[read] is { Container: { } container } known &&
        SameColumns(known.Setup.Columns, columns)
            ? new XElement(container.Name, container.Attributes())
            : null;

    private void Open(StringBuilder body, ColumnsDto columns, XElement? kept)
    {
        if (kept is not null)
        {
            _open = kept;
            body.Append(_out.Open(_open));
            return;
        }

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

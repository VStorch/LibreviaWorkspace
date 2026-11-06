using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>What the writers of one part share.</summary>
internal sealed class OdtWriteContext(OdtStyles styles, OdtAutomaticStyles automatic, OdtPictures pictures, Inventory inventory)
{
    private const string TablePrefix = "LvTable";
    private const string ImagePrefix = "LvImage";
    private const string EquationPrefix = "LvEquation";
    private const string FramePrefix = "LvFrame";
    private const string NotePrefix = "LvNote";
    private const string SectionPrefix = "LvSection";

    private readonly HashSet<string> _names = new(StringComparer.Ordinal);
    private readonly HashSet<string> _bookmarks = new(StringComparer.Ordinal);
    private readonly HashSet<string> _noteIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _reusedNoteIds = new(StringComparer.Ordinal);

    public OdtStyles Styles => styles;
    public OdtAutomaticStyles Automatic => automatic;
    public OdtPictures Pictures => pictures;
    public Inventory Inventory => inventory;

    /// <summary>The whole document's comments, notes and bookmarks; empty for a band.</summary>
    public OdtDocumentPlan Plan { get; init; } = OdtDocumentPlan.Empty;

    public OdtFormulas Formulas { get; init; } = new(_ => false);

    /// <summary>Inside a note, a band or a text box, where ODF allows no note.</summary>
    public int NotelessDepth { get; set; }

    /// <summary>In a header or footer, where <c>{n}</c> and <c>{total}</c> are page fields.</summary>
    public bool InBand { get; set; }

    /// <summary>The section of the block being written: notes may restart there.</summary>
    public int Section { get; set; }

    /// <summary>The list style of each numbering key written so far, so a list can continue another.</summary>
    public Dictionary<string, string> ListStyles { get; } = new(StringComparer.Ordinal);

    /// <summary>The table and frame names the part already has: a new one never repeats them.</summary>
    public void TakeNames(XElement root)
    {
        foreach (var element in root.Descendants())
        {
            if (element.Attr(Ns.Table + "name") is { } table) _names.Add(table);
            if (element.Attr(Ns.Draw + "name") is { } frame) _names.Add(frame);
            if (element.Name.LocalName == "note" && element.Attr(Ns.Text + "id") is { } note)
            {
                _names.Add(note);
                _noteIds.Add(note);
            }

            if (element.Name.LocalName == "section" && element.Attr(Ns.Text + "name") is { } section) _names.Add(section);
            if (element.Name.LocalName.StartsWith("bookmark", StringComparison.Ordinal) && element.Attr(Ns.Text + "name") is { } bookmark)
            {
                _bookmarks.Add(bookmark);
            }
        }
    }

    /// <summary>A name the document does not use yet; a repeated one is <c>false</c>, as LibreOffice refuses two.</summary>
    public bool TakeBookmark(string name) => _bookmarks.Add(name);

    /// <summary>The name of the table being rewritten: formulas and links point to it.</summary>
    public string? KeptTableName { get; set; }

    /// <summary>The cell padding of the table being rewritten.</summary>
    public string? KeptCellPadding { get; set; }

    public string NewTableName()
    {
        var kept = KeptTableName;
        KeptTableName = null;
        return kept ?? NewName(TablePrefix);
    }

    public string NewImageName() => NewName(ImagePrefix);

    public string NewEquationName() => NewName(EquationPrefix);

    public string NewFrameName() => NewName(FramePrefix);

    public string NewNoteId() => NewName(NotePrefix);

    /// <summary>
    /// The note's own id, so a <c>text:note-ref</c> in a kept block still finds it, when the original
    /// has it and the model holds it once (a kept block would be the other); otherwise a new one.
    /// </summary>
    public string NoteIdOf(string? nid) =>
        nid is not null && _noteIds.Contains(nid) && Plan.NoteHeldOnce(nid) && _reusedNoteIds.Add(nid) ? nid : NewNoteId();

    public string NewSectionName() => NewName(SectionPrefix);

    private string NewName(string prefix)
    {
        var number = 1;
        string name;
        while (!_names.Add(name = $"{prefix}{number}")) number++;
        return name;
    }

    /// <summary>The named style of the family for the editor's id, when the package has it.</summary>
    public string? NamedStyle(string family, string? id)
    {
        if (string.IsNullOrEmpty(id)) return null;
        var name = OdtWordStyles.OdfNameOf(family, id);
        return styles.IsNamed(family, name) ? name : null;
    }

    /// <summary>
    /// The heading style of the level, as the editor's cascade picks it: Word's id, or else any
    /// paragraph style named <c>heading N</c>.
    /// </summary>
    public string? HeadingStyle(int level) =>
        NamedStyle(OdtStyles.ParagraphFamily, OdtWordStyles.HeadingIdOf(level))
        ?? styles.Named.FirstOrDefault(style => style.Family == OdtStyles.ParagraphFamily &&
                                                string.Equals(style.DisplayName, OdtWordStyles.HeadingNameOf(level), StringComparison.OrdinalIgnoreCase))?.Name;

    /// <summary>The fields of the block being rewritten, in order: a written field takes the attributes of its own.</summary>
    public List<XElement> SourceFields { get; } = [];

    public XElement? TakeSourceField(string localName)
    {
        var found = SourceFields.FirstOrDefault(field => field.Name == Ns.Text + localName);
        if (found is not null) SourceFields.Remove(found);
        return found;
    }

    public void NoteLoss(string key, string? value = null) => inventory.NoteLoss(key, value);
}

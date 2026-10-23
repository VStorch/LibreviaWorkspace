using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// <b>Invisible</b>: stays in the file but does not show on screen. <b>Lost</b>: disappears on
/// save, and only what was in an edited block. <b>Structural</b>: invisible content that disappears
/// if the block anchoring it is edited; it decides read-only. Mixed together, they become a warning
/// people learn to ignore.
/// </summary>
/// <remarks>
/// Each item is a <see cref="Notice"/>: a key of <c>src/shared/i18n/catalog/sidecar.ts</c>, which
/// main shows in the interface language. The keys are constants because classification is done
/// <b>by key</b>: a key changed in one reader would stop matching and the document would open
/// editable.
/// </remarks>
public sealed class Inventory
{
    /// <summary>
    /// In headers, notes or text boxes, which the editor does not carry as nodes.
    /// </summary>
    public const string Comments = "sidecar.notice.comments";
    /// <summary>
    /// A revision the editor does not represent: cell, numbering, section or table.
    /// </summary>
    public const string StructureRevisions = "sidecar.notice.structureRevisions";

    /// <c>w:rPrChange</c> and <c>w:pPrChange</c>: in an edited paragraph, the range's is lost.
    public const string FormatRevisions = "sidecar.notice.formatRevisions";
    /// <summary>Inside a text box and in a draft older than notes (<c>BeforeNotes</c>).</summary>
    public const string Footnotes = "sidecar.notice.footnotes";
    public const string Endnotes = "sidecar.notice.endnotes";
    public const string Fields = "sidecar.notice.fields";
    public const string HeaderFields = "sidecar.notice.headerFields";
    /// <summary>
    /// A **frame** CSS cannot do: gradient, texture, shadow, 3D, rounded corner, or one inherited
    /// from a theme we do not resolve (see <see cref="ShapeLook"/>).
    /// </summary>
    public const string Shapes = "sidecar.notice.shapes";
    public const string ContentControls = "sidecar.notice.contentControls";
    /// <summary>
    /// An equation with a construct the screen does not draw: it shows locked, and the OMML goes
    /// back whole.
    /// </summary>
    public const string Equations = "sidecar.notice.equations";
    /// <summary>An element the readers do not know, named by its own tag.</summary>
    public const string Element = "sidecar.notice.element";

    private static readonly HashSet<string> StructuralKeys = new(StringComparer.Ordinal)
    {
        StructureRevisions, Fields, HeaderFields, ContentControls,
    };

    private static readonly Comparer<Notice> ByKeyThenValue = Comparer<Notice>.Create((left, right) =>
    {
        var byKey = string.CompareOrdinal(left.Key, right.Key);
        return byKey != 0 ? byKey : string.CompareOrdinal(left.Value, right.Value);
    });

    private readonly SortedSet<Notice> _invisible = new(ByKeyThenValue);
    private readonly SortedSet<Notice> _lost = new(ByKeyThenValue);
    private readonly SortedSet<Notice> _structural = new(ByKeyThenValue);

    [JsonPropertyName("invisible")]
    public IReadOnlyCollection<Notice> Invisible => _invisible;

    [JsonPropertyName("lost")]
    public IReadOnlyCollection<Notice> Lost => _lost;

    [JsonPropertyName("structural")]
    public IReadOnlyCollection<Notice> Structural => _structural;

    /// <param name="value">What the sentence quotes, such as a color or a sheet name.</param>
    public void NoteInvisible(string key, string? value = null)
    {
        var notice = new Notice(key, value);
        _invisible.Add(notice);
        if (StructuralKeys.Contains(key)) _structural.Add(notice);
    }

    /// <param name="value">What the sentence quotes, such as a color or a sheet name.</param>
    public void NoteLoss(string key, string? value = null) => _lost.Add(new Notice(key, value));

    /// <summary>
    /// Separate from <see cref="NoteInvisible"/>: an unknown lowercase name is noise, and an unknown
    /// uppercase one is named as it is.
    /// </summary>
    public void NoteInvisibleElement(string elementName)
    {
        var notice = Describe(elementName);
        if (notice is { } known) NoteInvisible(known.Key, known.Value);
    }

    /// <summary>
    /// Or <c>null</c> for noise: <c>w:bidi</c>, <c>w:textDirection</c> and <c>w:formProt</c> are in
    /// every section LibreOffice writes.
    /// </summary>
    private static Notice? Describe(string what) => what switch
    {
        "bidi" or "textDirection" or "formProt" or "docGrid" or "rPr" or "pPr" => null,
        "proofErr" or "lastRenderedPageBreak" or "bookmarkStart" or "bookmarkEnd" => null,
        "sectPr" or "tabs" or "spacing" or "ind" or "jc" or "widowControl" => null,

        "commentRangeStart" or "commentRangeEnd" => new Notice(Comments),
        "cellIns" or "cellDel" or "cellMerge" => new Notice(StructureRevisions),
        "footnoteReference" => new Notice(Footnotes),
        "endnoteReference" => new Notice(Endnotes),
        "fldChar" or "fldSimple" or "instrText" => new Notice(Fields),
        "pict" or "object" or "AlternateContent" => new Notice(Shapes),
        // Drawings and smart tags show as image and text: they are not structural.
        "drawing" => new Notice("sidecar.notice.drawings"),
        "smartTag" => new Notice("sidecar.notice.smartTags"),
        "sdt" or "sdtBlock" => new Notice(ContentControls),

        _ => what.Length > 0 && char.IsLower(what[0]) ? null : new Notice(Element, what),
    };

    [JsonIgnore]
    public bool IsEmpty => _invisible.Count == 0 && _lost.Count == 0;
}

/// <summary>
/// An item of the inventory: the catalog key and, for some, the value the sentence quotes.
/// </summary>
public readonly record struct Notice(
    [property: JsonPropertyName("key")] string Key,
    [property: JsonPropertyName("value")] string? Value = null);

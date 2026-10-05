using System.Text.Json.Serialization;

namespace Librevia.Format.Docx;

/// <summary>
/// <b>Invisible</b>: stays in the file but does not show on screen. <b>Lost</b>: disappears on
/// save, and only what was in an edited block. <b>Structural</b>: invisible content that disappears
/// if the block anchoring it is edited; it decides read-only. Mixed together, they become a warning
/// people learn to ignore.
/// </summary>
/// <remarks>
/// The labels are constants because classification is done **by label**: a sentence changed in one
/// reader would stop matching and the document would open editable.
/// </remarks>
public sealed class Inventory
{
    /// <summary>
    /// In headers, notes or text boxes, which the editor does not carry as nodes.
    /// </summary>
    public const string Comments = "comentários";
    /// <summary>
    /// A revision the editor does not represent: cell, numbering, section or table.
    /// </summary>
    public const string StructureRevisions = "revisões de estrutura";

    /// <c>w:rPrChange</c> and <c>w:pPrChange</c>: in an edited paragraph, the range's is lost.
    public const string FormatRevisions = "revisões de formatação";
    /// <summary>Inside a text box and in a draft older than notes (<c>BeforeNotes</c>).</summary>
    public const string Footnotes = "notas de rodapé";
    public const string Endnotes = "notas de fim";
    public const string Fields = "campos calculados (como sumário e número de página)";
    public const string HeaderFields = "campos calculados no cabeçalho";
    /// <summary>
    /// A **frame** CSS cannot do: gradient, texture, shadow, 3D, rounded corner, or one inherited
    /// from a theme we do not resolve (see <see cref="ShapeLook"/>).
    /// </summary>
    public const string Shapes = "moldura e preenchimento de formas";
    public const string ContentControls = "controles de conteúdo";
    /// <summary>
    /// An equation with a construct the screen does not draw: it shows locked, and the OMML goes
    /// back whole.
    /// </summary>
    public const string Equations = "equações";

    private static readonly HashSet<string> StructuralLabels = new(StringComparer.Ordinal)
    {
        StructureRevisions, Fields, HeaderFields, ContentControls,
    };

    private readonly SortedSet<string> _invisible = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _lost = new(StringComparer.Ordinal);
    private readonly SortedSet<string> _structural = new(StringComparer.Ordinal);

    [JsonPropertyName("invisible")]
    public IReadOnlyCollection<string> Invisible => _invisible;

    [JsonPropertyName("lost")]
    public IReadOnlyCollection<string> Lost => _lost;

    [JsonPropertyName("structural")]
    public IReadOnlyCollection<string> Structural => _structural;

    public void NoteInvisible(string message)
    {
        _invisible.Add(message);
        if (StructuralLabels.Contains(message)) _structural.Add(message);
    }

    public void NoteLoss(string message) => _lost.Add(message);

    /// <summary>
    /// Separate from <see cref="NoteInvisible"/>: the rule "an unknown lowercase name is noise"
    /// would swallow the Portuguese sentences.
    /// </summary>
    public void NoteInvisibleElement(string elementName)
    {
        var label = Describe(elementName);
        if (label is not null) NoteInvisible(label);
    }

    /// <summary>
    /// Or <c>null</c> for noise: <c>w:bidi</c>, <c>w:textDirection</c> and <c>w:formProt</c> are in
    /// every section LibreOffice writes.
    /// </summary>
    private static string? Describe(string what) => what switch
    {
        "bidi" or "textDirection" or "formProt" or "docGrid" or "rPr" or "pPr" => null,
        "proofErr" or "lastRenderedPageBreak" or "bookmarkStart" or "bookmarkEnd" => null,
        "sectPr" or "tabs" or "spacing" or "ind" or "jc" or "widowControl" => null,

        "commentRangeStart" or "commentRangeEnd" or "comentário" => Comments,
        "cellIns" or "cellDel" or "cellMerge" or "revisões de estrutura" => StructureRevisions,
        "footnoteReference" => Footnotes,
        "endnoteReference" => Endnotes,
        "fldChar" or "fldSimple" or "instrText" or "campo calculado" => Fields,
        "pict" or "object" or "AlternateContent" => Shapes,
        // Drawings and smart tags show as image and text: they are not structural.
        "drawing" => "desenhos",
        "smartTag" => "marcações inteligentes",
        "sdt" or "sdtBlock" => ContentControls,

        _ => what.Length > 0 && char.IsLower(what[0]) ? null : what,
    };

    [JsonIgnore]
    public bool IsEmpty => _invisible.Count == 0 && _lost.Count == 0;
}

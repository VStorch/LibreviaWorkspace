using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// What an ODT has and the editor does not represent. The document opens as a new one, so all of
/// it is <b>lost</b> on save: never silently. The keys mean the same as the DOCX reader's when the
/// thing is the same.
/// </summary>
internal static class OdtNotices
{
    public const string Revisions = "sidecar.notice.revisions";
    public const string Bookmarks = "sidecar.notice.bookmarks";
    public const string Sections = "sidecar.notice.sections";
    public const string TextBoxes = "sidecar.notice.textBoxes";
    public const string Forms = "sidecar.notice.forms";
    public const string Indexes = "sidecar.notice.indexes";
    public const string Drawings = "sidecar.notice.drawings";
    public const string ImagePosition = "sidecar.notice.imagePosition";
    public const string PaperSize = "sidecar.notice.paperSizeChanged";
    public const string HeaderVariants = "sidecar.notice.headerVariants";
    public const string BandFormatting = "sidecar.notice.bandFormatting";
    public const string ImageType = "sidecar.notice.imageType";
    public const string ImageWithoutContent = "sidecar.notice.imageWithoutContent";
    public const string ListNumberFormat = "sidecar.notice.listNumberFormat";

    /// <summary>Text fields: their current text stays, the calculation goes.</summary>
    private static readonly HashSet<string> TextFields = new(StringComparer.Ordinal)
    {
        "page-number", "page-count", "page-continuation", "date", "time", "author-name",
        "author-initials", "creator", "title", "subject", "description", "keywords", "initial-creator",
        "creation-date", "creation-time", "modification-date", "modification-time", "editing-cycles",
        "editing-duration", "file-name", "template-name", "chapter", "sequence", "sequence-ref",
        "reference-ref", "bookmark-ref", "note-ref", "variable-get", "variable-set", "variable-input",
        "user-field-get", "user-field-input", "expression", "text-input", "database-display",
        "database-name", "database-next", "database-row-number", "database-row-select",
        "sender-firstname", "sender-lastname", "sender-initials", "sender-title", "sender-position",
        "sender-email", "sender-phone-private", "sender-fax", "sender-company", "sender-phone-work",
        "sender-street", "sender-city", "sender-postal-code", "sender-country", "sender-state-or-province",
        "word-count", "character-count", "paragraph-count", "page-variable-get", "page-variable-set",
        "placeholder", "conditional-text", "hidden-text", "drop-down", "dde-connection", "execute-macro",
        "measure", "print-date", "print-time", "printed-by", "table-formula", "user-defined",
        "object-count", "image-count", "table-count", "sheet-name", "meta-field",
        "bibliography-mark",
    };

    private static readonly HashSet<string> BookmarkElements = new(StringComparer.Ordinal)
    {
        "bookmark", "bookmark-start", "bookmark-end", "reference-mark", "reference-mark-start", "reference-mark-end",
    };

    private static readonly HashSet<string> IndexMarks = new(StringComparer.Ordinal)
    {
        "toc-mark", "toc-mark-start", "toc-mark-end", "alphabetical-index-mark", "alphabetical-index-mark-start",
        "alphabetical-index-mark-end", "user-index-mark", "user-index-mark-start", "user-index-mark-end",
    };

    /// <summary>Their <c>text:index-body</c> stays as text.</summary>
    public static readonly HashSet<string> IndexBlocks = new(StringComparer.Ordinal)
    {
        "table-of-content", "alphabetical-index", "illustration-index", "table-index", "object-index",
        "user-index", "bibliography",
    };

    private static readonly HashSet<string> ChangeMarks = new(StringComparer.Ordinal)
    {
        "change", "change-start", "change-end",
    };

    public static bool IsTextField(XElement element) =>
        element.Name.Namespace == Ns.Text && TextFields.Contains(element.Name.LocalName);

    /// <summary>The key for an inline element that is dropped, or <c>null</c> when it is not one.</summary>
    public static string? DroppedInline(XElement element)
    {
        var name = element.Name.LocalName;
        if (element.Name.Namespace == Ns.Office) return name == "annotation" ? Inventory.Comments : null;
        if (element.Name.Namespace != Ns.Text) return null;
        if (name == "note") return element.Attr(Ns.Text + "note-class") == "endnote" ? Inventory.Endnotes : Inventory.Footnotes;
        if (BookmarkElements.Contains(name)) return Bookmarks;
        if (IndexMarks.Contains(name)) return Indexes;
        return ChangeMarks.Contains(name) ? Revisions : null;
    }

    /// <summary>A shape, a control or a 3D scene: drawn by LibreOffice, not by the editor.</summary>
    public static string ShapeKey(XElement element) =>
        element.Name.LocalName == "control" ? Forms : Drawings;

    public static void Element(Inventory inventory, XElement element) =>
        inventory.NoteLoss(Inventory.Element, Ns.Tag(element));
}

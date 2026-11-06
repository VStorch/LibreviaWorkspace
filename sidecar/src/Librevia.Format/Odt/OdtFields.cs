using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// The ODF text fields with a Word equivalent, which become the same <c>field</c> node the DOCX
/// reader gives: the instruction in Word's form, the current text as the result.
/// </summary>
internal static class OdtFields
{
    private const string PageNumber = "page-number";
    private const string CurrentPage = "current";

    /// <summary><c>text:creator</c> is the last author; <c>text:initial-creator</c>, Word's <c>AUTHOR</c>.</summary>
    private static readonly (string Element, string Kind)[] Pairs =
    [
        (PageNumber, "PAGE"),
        ("page-count", "NUMPAGES"),
        ("date", "DATE"),
        ("time", "TIME"),
        ("title", "TITLE"),
        ("subject", "SUBJECT"),
        ("initial-creator", "AUTHOR"),
        ("creator", "LASTSAVEDBY"),
        ("keywords", "KEYWORDS"),
        ("description", "COMMENTS"),
    ];

    /// <summary>
    /// <c>null</c> for one without an equivalent; a page number of the previous or next page, or
    /// shifted, is not Word's <c>PAGE</c>.
    /// </summary>
    public static string? InstructionOf(XElement element)
    {
        if (element.Name.Namespace != Ns.Text) return null;
        var name = element.Name.LocalName;
        if (name == PageNumber && !IsCurrentPage(element)) return null;
        var kind = Pairs.FirstOrDefault(pair => pair.Element == name).Kind;
        return kind is null ? null : $" {kind} ";
    }

    /// <summary>The ODF element's local name for a Word instruction, by its first word.</summary>
    public static string? ElementOf(string? instruction)
    {
        var word = (instruction ?? string.Empty).Trim().Split((char[]?)null, 2, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        if (word is null) return null;
        var element = Pairs.FirstOrDefault(pair => pair.Kind == word.ToUpperInvariant()).Element;
        return element;
    }

    /// <summary>A new page number, as LibreOffice writes one.</summary>
    public static XElement Create(string element)
    {
        var created = new XElement(Ns.Text + element);
        if (element == PageNumber) created.Add(new XAttribute(Ns.Text + "select-page", CurrentPage));
        return created;
    }

    private static bool IsCurrentPage(XElement element) =>
        (element.Attr(Ns.Text + "select-page") ?? CurrentPage) == CurrentPage &&
        (OdtLength.Integer(element.Attr(Ns.Text + "page-adjust")) ?? 0) == 0;
}

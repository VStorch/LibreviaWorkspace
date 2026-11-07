using System.Text.RegularExpressions;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>office:annotation</c> → the comment library and its anchors, as <see cref="BodyReader"/> gives
/// a DOCX comment: a named annotation whose <c>office:annotation-end</c> exists opens a range
/// (<c>commentStart</c>, then <c>commentEnd</c> at the end); any other is a point comment, only
/// <c>commentEnd</c>. A reply (<c>loext:parent-name</c>) has no anchor. The ids are positional in
/// <c>content.xml</c>, so reading the same file again gives the same ones.
/// </summary>
internal sealed partial class OdtCommentReader
{
    private static readonly HashSet<string> PlainElements = new(StringComparer.Ordinal)
    {
        "span", "s", "tab", "line-break", "soft-page-break",
    };

    private readonly HashSet<string> _ends;
    private readonly Dictionary<XElement, string> _idOf = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<string, string> _idByName = new(StringComparer.Ordinal);
    private readonly List<CommentDto> _library = [];
    private readonly HashSet<string> _ranged = new(StringComparer.Ordinal);

    /// <summary>
    /// Ids and names in a first pass: LibreOffice may write a reply before the comment it answers.
    /// </summary>
    public OdtCommentReader(XDocument content)
    {
        _ends = content.Descendants(Ns.Office + "annotation-end")
            .Select(end => end.Attr(Ns.Office + "name")).OfType<string>().ToHashSet(StringComparer.Ordinal);
        foreach (var annotation in content.Descendants(Ns.Office + "annotation"))
        {
            var id = _idOf.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            _idOf[annotation] = id;
            if (annotation.Attr(Ns.Office + "name") is { } name && _idByName.TryAdd(name, id)) Names[id] = name;
        }
    }

    public List<CommentDto>? Library => _library.Count == 0 ? null : _library;

    /// <summary>The <c>office:name</c> of each named comment, by id: saving keeps it.</summary>
    public Dictionary<string, string> Names { get; } = new(StringComparer.Ordinal);

    /// <param name="rich">Whether the comment holds more than plain text: lost if its block is rewritten.</param>
    public Node? Annotation(XElement annotation, out bool rich)
    {
        var id = _idOf[annotation];
        var name = annotation.Attr(Ns.Office + "name");
        var parentId = annotation.Attr(Ns.Loext + "parent-name") is { } parent ? _idByName.GetValueOrDefault(parent) : null;
        if (parentId == id) parentId = null;

        rich = IsRich(annotation);
        _library.Add(new CommentDto(
            id,
            annotation.Element(Ns.Dc + "creator")?.Value ?? string.Empty,
            annotation.Element(Ns.Dc + "date")?.Value ?? string.Empty,
            [.. annotation.Elements(Ns.Text + "p").Select(TextOf)],
            annotation.Attr(Ns.Loext + "resolved") == "true",
            ParentId: parentId,
            Initials: annotation.Element(Ns.Meta + "creator-initials")?.Value is { Length: > 0 } initials ? initials : null,
            Rich: rich));

        if (parentId is not null) return null;
        if (name is null || !_ends.Contains(name)) return Node.Of("commentEnd").With("cid", id);
        _ranged.Add(id);
        return Node.Of("commentStart").With("cid", id);
    }

    /// <summary>The end of a range opened by a comment that is not a reply.</summary>
    public Node? End(XElement end) =>
        end.Attr(Ns.Office + "name") is { } name && _idByName.TryGetValue(name, out var id) && _ranged.Contains(id)
            ? Node.Of("commentEnd").With("cid", id)
            : null;

    /// <summary>Anything but paragraphs of text, spans and white space: a field, a link, a picture, a list.</summary>
    private static bool IsRich(XElement annotation) =>
        annotation.Elements().Any(child => child.Name.Namespace == Ns.Text && child.Name != Ns.Text + "p") ||
        annotation.Elements(Ns.Text + "p").Descendants().Any(element => element.Name.Namespace != Ns.Text || !PlainElements.Contains(element.Name.LocalName));

    /// <summary>As the writer gives it: collapsed white space, real spaces, tabs and line breaks.</summary>
    private static string TextOf(XElement paragraph) => string.Concat(paragraph.DescendantNodes().Select(node => node switch
    {
        XText text => WhiteSpace().Replace(text.Value, " "),
        XElement { Name.LocalName: "s" } space => new string(' ', Math.Clamp(OdtLength.Integer(space.Attr(Ns.Text + "c")) ?? 1, 1, OdtInlineReader.MaxSpaces)),
        XElement { Name.LocalName: "tab" } => "\t",
        XElement { Name.LocalName: "line-break" } => "\n",
        _ => string.Empty,
    }));

    [GeneratedRegex(@"[ \t\r\n]+")]
    private static partial Regex WhiteSpace();
}

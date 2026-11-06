using System.Text.RegularExpressions;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inline nodes that are more than text: notes with their citation, comments as annotations
/// (the first one spans up to its end anchor, the replies follow), bookmarks and equations as
/// formula objects drawn like a character.
/// </summary>
internal sealed partial class OdtInlineObjects(OdtWriteContext context, OdtBlockWriter blocks)
{
    private const double MinFormulaWidthMm = 3;
    private const double MaxFormulaWidthMm = 170;
    private const double FormulaMmPerCharacter = 2.2;
    private const double FormulaBaseHeightMm = 5;
    private const double FormulaMmPerLevel = 3.5;
    private const double MaxFormulaHeightMm = 120;

    public static readonly HashSet<string> Types = new(StringComparer.Ordinal)
    {
        "noteRef", "commentStart", "commentEnd", "bookmarkStart", "bookmarkEnd", "math",
    };

    public IEnumerable<XNode> Write(Node node) => node.Type switch
    {
        "noteRef" => Note(node),
        "commentStart" => CommentStart(node),
        "commentEnd" => CommentEnd(node),
        "bookmarkStart" => BookmarkStart(node),
        "bookmarkEnd" => BookmarkEnd(node),
        _ => Formula(node),
    };

    private IEnumerable<XNode> Note(Node node)
    {
        var endnote = Attr.String(node, "kind") == OdtNoteCounter.Endnote;
        if (context.NotelessDepth > 0)
        {
            context.NoteLoss(endnote ? Inventory.Endnotes : Inventory.Footnotes);
            return [];
        }

        var mark = Attr.String(node, "mark");
        var label = context.Plan.Notes.Next(endnote, context.Section, mark);
        var citation = new XElement(Ns.Text + "note-citation", label);
        if (!string.IsNullOrEmpty(mark)) citation.Add(new XAttribute(Ns.Text + "label", mark));

        context.NotelessDepth++;
        var body = new XElement(Ns.Text + "note-body", blocks.WriteAll(node.Content ?? []));
        context.NotelessDepth--;
        if (!body.HasElements) body.Add(new XElement(Ns.Text + "p"));

        return [new XElement(
            Ns.Text + "note",
            new XAttribute(Ns.Text + "id", context.NoteIdOf(Attr.String(node, "nid"))),
            new XAttribute(Ns.Text + "note-class", endnote ? OdtNoteCounter.Endnote : OdtNoteCounter.Footnote),
            citation,
            body)];
    }

    private IEnumerable<XNode> CommentStart(Node node)
    {
        var id = Attr.String(node, "cid") ?? string.Empty;
        var ranged = context.Plan.CommentEnds(id);
        var written = Thread(id, ranged).ToList();
        if (ranged && written.Count > 0) context.Plan.OpenComments.Add(id);
        return written;
    }

    /// <summary>
    /// The root, named when a range or a reply points to it, and its replies right after it, each
    /// pointing to it by <c>loext:parent-name</c>, as LibreOffice threads them.
    /// </summary>
    private IEnumerable<XElement> Thread(string id, bool ranged)
    {
        var plan = context.Plan;
        var root = plan.Comments.FirstOrDefault(comment => comment.Id == id);
        if (root is null || !plan.WrittenComments.Add(id))
        {
            context.NoteLoss(Inventory.Comments);
            return [];
        }

        var replies = plan.Comments.Where(comment => comment.ParentId == id).ToList();
        var name = ranged || replies.Count > 0 ? plan.NameOf(id) : null;
        return [Annotation(root, root.Done, name, null), .. replies.Select(reply => Annotation(reply, root.Done, null, name))];
    }

    private static XElement Annotation(CommentDto comment, bool resolved, string? name, string? parentName)
    {
        var annotation = new XElement(Ns.Office + "annotation");
        if (name is not null) annotation.Add(new XAttribute(Ns.Office + "name", name));
        if (parentName is not null) annotation.Add(new XAttribute(Ns.Loext + "parent-name", parentName));
        if (resolved) annotation.Add(new XAttribute(Ns.Loext + "resolved", "true"));
        annotation.Add(new XElement(Ns.Dc + "creator", comment.Author));
        if (IsoDate().IsMatch(comment.Date)) annotation.Add(new XElement(Ns.Dc + "date", comment.Date));
        if (comment.Initials is { Length: > 0 } initials) annotation.Add(new XElement(Ns.Meta + "creator-initials", initials));
        var paragraphs = comment.Paragraphs.Count == 0 ? [string.Empty] : comment.Paragraphs;
        annotation.Add(paragraphs.Select(text => new XElement(Ns.Text + "p", OdtInlineWriter.PlainText(text))));
        return annotation;
    }

    /// <summary>
    /// Closes the range opened here, or one whose start stayed as it was in the original; without a
    /// start anywhere, the comment is a point and is written here.
    /// </summary>
    private IEnumerable<XNode> CommentEnd(Node node)
    {
        var id = Attr.String(node, "cid") ?? string.Empty;
        var plan = context.Plan;
        if (!plan.CommentStarts(id)) return Thread(id, ranged: false);
        var name = plan.OpenComments.Contains(id) ? plan.NameOf(id) : plan.KeptNameOf(id);
        return name is null ? [] : [new XElement(Ns.Office + "annotation-end", new XAttribute(Ns.Office + "name", name))];
    }

    private IEnumerable<XNode> BookmarkStart(Node node)
    {
        if (Attr.String(node, "name") is not { Length: > 0 } name || !context.TakeBookmark(name))
        {
            context.NoteLoss(OdtNotices.Bookmarks);
            return [];
        }

        var bid = Attr.String(node, "bid") ?? string.Empty;
        if (!context.Plan.BookmarkEnds(bid)) return [new XElement(Ns.Text + "bookmark", new XAttribute(Ns.Text + "name", name))];
        context.Plan.BookmarkNames[bid] = name;
        return [new XElement(Ns.Text + "bookmark-start", new XAttribute(Ns.Text + "name", name))];
    }

    private IEnumerable<XNode> BookmarkEnd(Node node) =>
        context.Plan.BookmarkNames.TryGetValue(Attr.String(node, "bid") ?? string.Empty, out var name)
            ? [new XElement(Ns.Text + "bookmark-end", new XAttribute(Ns.Text + "name", name))]
            : [];

    /// <summary>One that fails the filter leaves no object and no text, and is declared.</summary>
    private IEnumerable<XNode> Formula(Node node)
    {
        if (MathMlFilter.Sanitize(Attr.String(node, "mathml")) is not { } math)
        {
            context.NoteLoss(Inventory.Equations);
            return [];
        }

        var folder = context.Formulas.Add(math, Attr.Bool(node, "display"));
        var (width, height) = FormulaSizeMm(math);
        var properties = new XElement(
            Ns.Style + "graphic-properties",
            new XAttribute(Ns.Style + "vertical-pos", "middle"),
            new XAttribute(Ns.Style + "vertical-rel", "text"),
            new XAttribute(Ns.Fo + "padding", "0mm"),
            new XAttribute(Ns.Fo + "border", "none"));
        var frame = new XElement(
            Ns.Draw + "frame",
            new XAttribute(Ns.Draw + "style-name", context.Automatic.Style(OdtAutomaticStyles.Graphic, OdtAutomaticStyles.GraphicFamily, null, properties)),
            new XAttribute(Ns.Draw + "name", context.NewEquationName()),
            new XAttribute(Ns.Text + "anchor-type", "as-char"),
            new XAttribute(Ns.Svg + "width", OdtFormatWriter.Mm(width)!),
            new XAttribute(Ns.Svg + "height", OdtFormatWriter.Mm(height)!),
            new XAttribute(Ns.Draw + "z-index", 0),
            new XElement(
                Ns.Draw + "object",
                new XAttribute(Ns.XLink + "href", "./" + folder),
                new XAttribute(Ns.XLink + "type", "simple"),
                new XAttribute(Ns.XLink + "show", "embed"),
                new XAttribute(Ns.XLink + "actuate", "onLoad")));
        if (Attr.String(node, "latex")?.Trim() is { Length: > 0 } latex) frame.Add(new XElement(Ns.Svg + "desc", latex));
        return [frame];
    }

    /// <summary>
    /// Estimated: whoever draws the formula adjusts it, but the frame needs a size so the line does
    /// not jump when nobody draws. Width comes from the text; height from how much the equation stacks.
    /// </summary>
    public static (double Width, double Height) FormulaSizeMm(XElement math)
    {
        var characters = MathMlFilter.TextOf(math).EnumerateRunes().Count();
        var width = Math.Clamp((characters * FormulaMmPerCharacter) + 1, MinFormulaWidthMm, MaxFormulaWidthMm);
        var height = Math.Min(MaxFormulaHeightMm, FormulaBaseHeightMm + (MathMlFilter.StackedLevels(math) * FormulaMmPerLevel));
        return (width, height);
    }

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}")]
    private static partial Regex IsoDate();
}

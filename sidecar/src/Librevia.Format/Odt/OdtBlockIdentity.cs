using System.Text;
using System.Text.Json;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// What decides whether a block comes back as it was. The fingerprint leaves out a note's body and
/// a comment's text, which the DOCX writer compares in their own parts; in an ODT both live inside
/// the block, so they count here.
/// </summary>
internal static class OdtBlockIdentity
{
    public static string Of(Node node, IReadOnlyList<CommentDto> comments)
    {
        var copy = JsonSerializer.Deserialize<Node>(JsonSerializer.Serialize(node, DocxJson.Options), DocxJson.Options)!;
        WithoutZeroIndent(copy);
        foreach (var math in new[] { copy }.Concat(Descendants(copy)).Where(inner => inner.Type == "math")) KeptOfFormula(math);
        var identity = new StringBuilder(copy.Fingerprint());
        foreach (var inner in Descendants(copy))
        {
            if (inner.Type == "noteRef")
            {
                foreach (var block in inner.Content ?? []) identity.Append('|').Append(block.Fingerprint());
            }
            else if (inner.Type is "commentStart" or "commentEnd")
            {
                identity.Append('|').Append(ThreadOf(comments, Attr.String(inner, "cid") ?? string.Empty));
            }
        }

        return identity.ToString();
    }

    /// <summary>In text order, note bodies included.</summary>
    public static IEnumerable<Node> Descendants(Node node) =>
        (node.Content ?? []).SelectMany(child => new[] { child }.Concat(Descendants(child)));

    /// <summary>The comment and its replies, as the pane edits them.</summary>
    private static string ThreadOf(IReadOnlyList<CommentDto> comments, string id) => JsonSerializer.Serialize(
        comments.Where(comment => comment.Id == id || comment.ParentId == id), DocxJson.Options);

    /// <summary>
    /// What the formula object keeps: its filtered MathML and mode. The LaTeX the editor derives,
    /// and its own flags, are not in the file.
    /// </summary>
    private static void KeptOfFormula(Node math)
    {
        var markup = MathMlFilter.Sanitize(Attr.String(math, "mathml"))?.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        math.Attrs = new() { ["mathml"] = markup, ["display"] = Attr.Bool(math, "display") };
    }

    /// <summary>
    /// The editor gives every paragraph an indent level, zero when it has none; an ODT states its
    /// indent in millimetres.
    /// </summary>
    private static void WithoutZeroIndent(Node node)
    {
        if (node.Attrs?.TryGetValue("indent", out var indent) == true && indent?.ToJsonString() == "0") node.Attrs.Remove("indent");
        foreach (var child in node.Content ?? []) WithoutZeroIndent(child);
    }
}

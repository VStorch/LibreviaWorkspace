using System.Text;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A comment as the pane shows it. Outside the nodes, like styles: the body lives in
/// <c>word/comments.xml</c>, and the paragraph only carries the anchor ends.
/// </summary>
public sealed record CommentDto(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("author")] string Author,
    [property: JsonPropertyName("date")] string Date,
    [property: JsonPropertyName("paragraphs")] List<string> Paragraphs,
    [property: JsonPropertyName("done")] bool Done,
    // `w15:paraIdParent` translated to the parent's id; absent on the one opening the thread.
    [property: JsonPropertyName("parentId")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? ParentId = null,
    [property: JsonPropertyName("initials")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Initials = null,
    // The last paragraph's `w14:paraId`, through which `commentsExtended.xml` links replies and
    // "resolved".
    [property: JsonPropertyName("paraId")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? ParaId = null,
    // Formatting, an image, a field or a link plain text does not show: it stays in the file, and
    // the pane warns.
    [property: JsonPropertyName("rich")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Rich = false);

/// <c>word/comments.xml</c> and <c>word/commentsExtended.xml</c> → the pane's comments. A malformed
/// one comes out with whatever could be read.
public static class CommentsReader
{
    public static List<CommentDto>? Read(MainDocumentPart part)
    {
        var comments = part.WordprocessingCommentsPart?.Comments?.Elements<Comment>().ToList() ?? [];
        if (comments.Count == 0) return null;

        var threads = ThreadsOf(part);
        var result = new List<CommentDto>();
        foreach (var comment in comments)
        {
            if (comment.Id?.Value is not { Length: > 0 } id) continue;

            var paragraphs = comment.Elements<Paragraph>().ToList();
            var paraId = paragraphs.LastOrDefault()?.ParagraphId?.Value;
            var extended = paraId is null ? null : threads.GetValueOrDefault(paraId);

            result.Add(new CommentDto(
                id,
                comment.Author?.Value ?? string.Empty,
                comment.Date?.InnerText ?? string.Empty,
                paragraphs.Select(TextOf).ToList(),
                extended?.Done ?? false,
                ParentId: extended?.Parent,
                Initials: comment.Initials?.Value is { Length: > 0 } initials ? initials : null,
                ParaId: paraId,
                Rich: IsRich(comment)));
        }

        return result.Count == 0 ? null : result;
    }

    /// <summary>
    /// Each comment's replies, in file order. A reply does not become a node: its anchor is the
    /// parent's, and saving returns it next to it.
    /// </summary>
    public static Dictionary<string, List<string>> RepliesOf(MainDocumentPart part)
    {
        var replies = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (var comment in Read(part) ?? [])
        {
            if (comment.ParentId is not { } parent) continue;
            if (!replies.TryGetValue(parent, out var list)) replies[parent] = list = [];
            list.Add(comment.Id);
        }

        return replies;
    }

    internal sealed record ThreadEntry(bool Done, string? Parent);

    /// <summary>
    /// `w15:done` and `w15:paraIdParent`, by the comment's `paraId`, with the parent already
    /// translated to an id.
    /// </summary>
    internal static Dictionary<string, ThreadEntry> ThreadsOf(MainDocumentPart part)
    {
        var entries = part.WordprocessingCommentsExPart?.CommentsEx?.Elements<W15.CommentEx>().ToList() ?? [];
        var result = new Dictionary<string, ThreadEntry>(StringComparer.OrdinalIgnoreCase);
        if (entries.Count == 0) return result;

        var idByParaId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var comment in part.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? [])
        {
            if (comment.Id?.Value is { Length: > 0 } id &&
                comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value is { } paraId)
            {
                idByParaId.TryAdd(paraId, id);
            }
        }

        foreach (var entry in entries)
        {
            if (entry.ParaId?.Value is not { } paraId) continue;
            var parent = entry.ParaIdParent?.Value is { } parentParaId
                ? idByParaId.GetValueOrDefault(parentParaId)
                : null;
            result.TryAdd(paraId, new ThreadEntry(entry.Done?.Value == true, parent));
        }

        return result;
    }

    internal static string TextOf(Paragraph paragraph)
    {
        var text = new StringBuilder();
        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case Text piece: text.Append(piece.Text); break;
                case TabChar: text.Append('\t'); break;
                case Break: text.Append('\n'); break;
            }
        }

        return text.ToString();
    }

    private static bool IsRich(Comment comment) =>
        comment.Descendants().Any(element => element is
            DocumentFormat.OpenXml.Wordprocessing.Drawing or Picture or FieldChar or SimpleField or Hyperlink or Table
            or Bold or Italic or Underline or Color or Strike or Highlight or VerticalTextAlignment);
}

using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// What an inline node needs from the whole document: the comment library and which comments and
/// bookmarks close somewhere (a range, not a point), and the note counter.
/// </summary>
internal sealed class OdtDocumentPlan
{
    public static readonly OdtDocumentPlan Empty = new(null, null, null);

    private readonly HashSet<string> _commentEnds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _bookmarkEnds = new(StringComparer.Ordinal);

    public OdtDocumentPlan(Node? doc, List<CommentDto>? comments, NotesDto? notes)
    {
        Comments = comments ?? [];
        Notes = new OdtNoteCounter(notes);
        if (doc is not null) Scan(doc);
    }

    public IReadOnlyList<CommentDto> Comments { get; }

    public OdtNoteCounter Notes { get; }

    /// <summary>The comments whose range is open: their end anchor closes it.</summary>
    public HashSet<string> OpenComments { get; } = new(StringComparer.Ordinal);

    /// <summary>The name each written bookmark start took, by its <c>bid</c>.</summary>
    public Dictionary<string, string> BookmarkNames { get; } = new(StringComparer.Ordinal);

    public bool CommentEnds(string id) => _commentEnds.Contains(id);

    public bool BookmarkEnds(string bid) => _bookmarkEnds.Contains(bid);

    private void Scan(Node node)
    {
        if (node.Type == "commentEnd" && Attr.String(node, "cid") is { } cid) _commentEnds.Add(cid);
        if (node.Type == "bookmarkEnd" && Attr.String(node, "bid") is { } bid) _bookmarkEnds.Add(bid);
        foreach (var child in node.Content ?? []) Scan(child);
    }
}

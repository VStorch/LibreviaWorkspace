using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// What an inline node needs from the whole document: the comment library and which comments and
/// bookmarks close somewhere (a range, not a point), and the note counter.
/// </summary>
internal sealed class OdtDocumentPlan
{
    public const string AnnotationPrefix = "__Annotation__";

    public static readonly OdtDocumentPlan Empty = new(null, null, null);

    private readonly HashSet<string> _commentStarts = new(StringComparer.Ordinal);
    private readonly HashSet<string> _commentEnds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _bookmarkEnds = new(StringComparer.Ordinal);
    private readonly IReadOnlyDictionary<string, string> _keptNames;
    private readonly HashSet<string> _annotationNames;
    private readonly Dictionary<string, string> _names = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _noteIds = new(StringComparer.Ordinal);

    /// <param name="keptNames">The <c>office:name</c> each comment of the original has, by id.</param>
    /// <param name="original">The original <c>content.xml</c>: a new name never repeats one of its annotations'.</param>
    public OdtDocumentPlan(
        Node? doc, List<CommentDto>? comments, NotesDto? notes,
        IReadOnlyDictionary<string, string>? keptNames = null, XDocument? original = null)
    {
        Comments = comments ?? [];
        Notes = new OdtNoteCounter(notes);
        _keptNames = keptNames ?? new Dictionary<string, string>();
        _annotationNames = (original?.Descendants(Ns.Office + "annotation") ?? [])
            .Select(annotation => annotation.Attr(Ns.Office + "name")).OfType<string>().ToHashSet(StringComparer.Ordinal);
        if (doc is not null) Scan(doc);
    }

    public IReadOnlyList<CommentDto> Comments { get; }

    public OdtNoteCounter Notes { get; }

    /// <summary>The comments whose range is open: their end anchor closes it.</summary>
    public HashSet<string> OpenComments { get; } = new(StringComparer.Ordinal);

    /// <summary>The threads written so far: each goes to the file once.</summary>
    public HashSet<string> WrittenComments { get; } = new(StringComparer.Ordinal);

    /// <summary>The name each written bookmark start took, by its <c>bid</c>.</summary>
    public Dictionary<string, string> BookmarkNames { get; } = new(StringComparer.Ordinal);

    public bool CommentEnds(string id) => _commentEnds.Contains(id);

    /// <summary>Without a start the comment is a point, written where its end is.</summary>
    public bool CommentStarts(string id) => _commentStarts.Contains(id);

    /// <summary>The name a kept anchor of the comment already uses, if any.</summary>
    public string? KeptNameOf(string id) => _keptNames.GetValueOrDefault(id);

    /// <summary>
    /// The original's name, so a range whose other end stays as it was still closes; a new comment
    /// takes one no annotation of the original has.
    /// </summary>
    public string NameOf(string id)
    {
        if (_keptNames.TryGetValue(id, out var kept)) return kept;
        if (_names.TryGetValue(id, out var given)) return given;
        var name = AnnotationPrefix + id;
        for (var suffix = 1; _annotationNames.Contains(name); suffix++) name = $"{AnnotationPrefix}{id}_{suffix}";
        _annotationNames.Add(name);
        return _names[id] = name;
    }

    public bool BookmarkEnds(string bid) => _bookmarkEnds.Contains(bid);

    public bool NoteHeldOnce(string nid) => _noteIds.GetValueOrDefault(nid) == 1;

    private void Scan(Node node)
    {
        if (node.Type == "commentStart" && Attr.String(node, "cid") is { } start) _commentStarts.Add(start);
        if (node.Type == "commentEnd" && Attr.String(node, "cid") is { } cid) _commentEnds.Add(cid);
        if (node.Type == "bookmarkEnd" && Attr.String(node, "bid") is { } bid) _bookmarkEnds.Add(bid);
        if (node.Type == "noteRef" && Attr.String(node, "nid") is { } nid) _noteIds[nid] = _noteIds.GetValueOrDefault(nid) + 1;
        foreach (var child in node.Content ?? []) Scan(child);
    }
}

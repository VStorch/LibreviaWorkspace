using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inverse of the revisions <see cref="OdtInlineReader"/> reads: each run of the same insertion
/// between <c>text:change-start</c> and <c>text:change-end</c>; each run of the same deletion, deleted
/// paragraph marks included, as one <c>text:change</c> whose region holds what it took away.
/// </summary>
internal sealed partial class OdtInlineWriter
{
    private OdtOpenChange? _open;
    private DeletedRun? _deleted;
    private List<XNode> _tracked = [];

    /// <summary>What a deletion took away, one paragraph per deleted paragraph mark.</summary>
    private sealed class DeletedRun(Mark mark)
    {
        public Mark Mark => mark;
        public List<XElement> Paragraphs { get; } = [new XElement(Ns.Text + "p")];
        public bool AfterSpace { get; set; } = true;
    }

    /// <summary>
    /// The content of paragraphs joined by deleted paragraph marks, which ODF writes as one
    /// paragraph; <paramref name="shellOf"/> gives the empty paragraph a joined one is stored as.
    /// </summary>
    public List<XNode> Write(IReadOnlyList<Node> chain, Func<Node, XElement> shellOf)
    {
        if (context.Changes is not { } changes) return Write(chain.SelectMany(paragraph => paragraph.Content ?? []));

        _afterSpace = true;
        _tracked = [];
        _open = changes.TakePending();
        for (var index = 0; index < chain.Count; index++)
        {
            if (index > 0) Delete(ParagraphMark(chain[index - 1])!, null, shellOf(chain[index]));
            foreach (var node in chain[index].Content ?? []) Tracked(node);
        }

        FlushDeletion();
        EndParagraph(ParagraphMark(chain[^1]));
        return _tracked;
    }

    private static Mark? ParagraphMark(Node paragraph) => OdtTrackedChanges.MarkOfBlock(Attr.Node(paragraph, "markRevision"));

    /// <summary>What the reader gives a revision mark: text, line breaks, fields, images and formulas.</summary>
    private static bool CarriesRevision(Node node) => node.Type is "text" or "hardBreak" or "field" or "image" or "math";

    private static Mark? RevisionOf(Node node, string type) =>
        CarriesRevision(node) ? node.Marks?.FirstOrDefault(mark => mark.Type == type) : null;

    private void Tracked(Node node)
    {
        if (RevisionOf(node, Revisions.Deletion) is { } deletion)
        {
            Delete(deletion, node, null);
            return;
        }

        FlushDeletion();
        if (CarriesRevision(node)) Insert(RevisionOf(node, Revisions.Insertion));
        _tracked.AddRange(Inline(node));
    }

    /// <summary>An insertion that was deleted is a deletion: ODF has no insertion inside one.</summary>
    private void Delete(Mark deletion, Node? node, XElement? joined)
    {
        if (_deleted is not null && !OdtTrackedChanges.Same(_deleted.Mark, deletion)) FlushDeletion();
        _deleted ??= new DeletedRun(deletion);
        if (joined is not null)
        {
            _deleted.Paragraphs.Add(joined);
            _deleted.AfterSpace = true;
        }

        if (node is null) return;
        if (RevisionOf(node, Revisions.Insertion) is not null) context.NoteLoss(OdtNotices.Revisions);
        var visible = _afterSpace;
        _afterSpace = _deleted.AfterSpace;
        _deleted.Paragraphs[^1].Add(Inline(node));
        _deleted.AfterSpace = _afterSpace;
        _afterSpace = visible;
    }

    private void FlushDeletion()
    {
        if (_deleted is null) return;
        _tracked.Add(context.Changes!.Point(_deleted.Mark, _deleted.Paragraphs));
        _deleted = null;
    }

    private void Insert(Mark? insertion)
    {
        if (_open is not null && insertion is not null && OdtTrackedChanges.Same(_open.Mark, insertion)) return;
        if (_open is not null) _tracked.Add(OdtChangeWriter.Close(_open));
        _open = insertion is null ? null : context.Changes!.Open(insertion);
        if (_open is not null) _tracked.Add(_open.Start);
    }

    /// <summary>An inserted paragraph mark leaves the range open; a deleted one that joins nothing stays.</summary>
    private void EndParagraph(Mark? paragraphMark)
    {
        if (paragraphMark?.Type == Revisions.Insertion)
        {
            Insert(paragraphMark);
            context.Changes!.Carry(_open!);
            _open = null;
            return;
        }

        if (paragraphMark is not null) context.NoteLoss(OdtNotices.Revisions);
        Insert(null);
    }
}

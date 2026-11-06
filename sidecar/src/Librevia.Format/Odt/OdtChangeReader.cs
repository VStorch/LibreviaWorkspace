using System.Text.Json.Nodes;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The revisions open while the text is read. An insertion range may cross paragraphs: what it
/// covers takes its mark, and each paragraph end inside it is an inserted paragraph mark. A note is
/// a text of its own, read with nothing open (<see cref="Suspend"/>).
/// </summary>
internal sealed class OdtChangeReader(OdtTrackedChanges changes)
{
    private List<OdtChangedRegion> _open = [];
    private Mark? _deletion;

    public OdtTrackedChanges Source => changes;

    public List<Mark>? Marks()
    {
        if (_deletion is not null) return [_deletion];
        return _open.Count > 0 ? [_open[^1].Mark] : null;
    }

    /// <summary>False when the range is not an insertion this reader represents.</summary>
    public bool Start(string? id)
    {
        if (_deletion is not null || changes.Find(id) is not { Kind: OdtChangedRegion.Insertion } region) return false;
        _open.Add(region);
        return true;
    }

    public void End(string? id)
    {
        if (changes.Find(id) is { } region) _open.Remove(region);
    }

    /// <summary>The deletion the point brings back; none inside another one.</summary>
    public OdtChangedRegion? Deletion(string? id) =>
        _deletion is null && changes.Find(id) is { Kind: OdtChangedRegion.Deletion } region ? region : null;

    public void BeginDeletion(Mark mark) => _deletion = mark;

    public void EndDeletion() => _deletion = null;

    /// <summary>The revision of a paragraph that ends here: inserted while a range is open.</summary>
    public JsonObject? ParagraphEnd() =>
        _deletion is null && _open.Count > 0 ? OdtTrackedChanges.BlockRevisionOf(_open[^1].Mark) : null;

    public List<OdtChangedRegion> Suspend()
    {
        var open = _open;
        _open = [];
        return open;
    }

    public void Resume(List<OdtChangedRegion> open) => _open = open;
}

using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// <c>w:vMerge</c> as the editor's <c>rowspan</c>: a restart cell spans itself plus every
/// continuation below it in the same grid column with the same <c>w:gridSpan</c>. A continuation
/// with nothing open above it stays an ordinary cell.
/// </summary>
internal sealed class VerticalMerges
{
    private readonly Dictionary<TableCell, int> _spans = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<TableCell> _absorbed = new(ReferenceEqualityComparer.Instance);

    private VerticalMerges()
    {
    }

    /// <summary>The reading from before merges: every cell its own, every <c>w:vMerge</c> left alone.</summary>
    public static readonly VerticalMerges None = new();

    public static VerticalMerges Of(Table? table)
    {
        var merges = new VerticalMerges();
        if (table is null) return merges;

        var open = new Dictionary<int, (TableCell Owner, int Span)>();
        foreach (var row in table.Elements<TableRow>())
        {
            open = merges.ReadRow(row, open);
        }

        return merges;
    }

    /// <summary>How many rows the cell spans; 1 for a continuation or an unmerged cell.</summary>
    public int RowSpanOf(TableCell cell) => _spans.GetValueOrDefault(cell, 1);

    /// <summary>A continuation folded into the cell above it, absent from the editor's model.</summary>
    public bool IsAbsorbed(TableCell cell) => _absorbed.Contains(cell);

    /// <summary>Takes part in a merge, as its first cell or as a continuation.</summary>
    public bool IsMerged(TableCell cell) => RowSpanOf(cell) > 1 || IsAbsorbed(cell);

    /// <summary><c>w:gridBefore</c>: the grid columns the row leaves empty before its first cell.</summary>
    public static int GridBeforeOf(TableRow? row) =>
        row?.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value is > 0 and var before ? before : 0;

    public static int GridSpanOf(TableCell cell) =>
        cell.TableCellProperties?.GridSpan?.Val?.Value is > 1 and var span ? span : 1;

    private Dictionary<int, (TableCell Owner, int Span)> ReadRow(
        TableRow row,
        Dictionary<int, (TableCell Owner, int Span)> open)
    {
        var next = new Dictionary<int, (TableCell Owner, int Span)>();
        var column = GridBeforeOf(row);

        foreach (var cell in row.Elements<TableCell>())
        {
            var span = GridSpanOf(cell);
            var merge = cell.TableCellProperties?.VerticalMerge;
            var restarts = merge?.Val?.Value == MergedCellValues.Restart;

            if (merge is not null && !restarts && open.TryGetValue(column, out var above) && above.Span == span)
            {
                _spans[above.Owner] = RowSpanOf(above.Owner) + 1;
                _absorbed.Add(cell);
                next[column] = above;
            }
            else if (restarts)
            {
                next[column] = (cell, span);
            }

            column += span;
        }

        return next;
    }
}

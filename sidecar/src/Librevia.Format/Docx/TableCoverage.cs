namespace Librevia.Format.Docx;

/// <summary>A grid position in a row: an editor cell, or the continuation of one above.</summary>
internal readonly record struct TableSlot(Node? Cell, int Span)
{
    public bool IsContinuation => Cell is null;
}

/// <summary>
/// TableKit leaves out of a row the columns a <c>rowspan</c> above covers; <c>w:tr</c> needs a
/// <c>w:vMerge</c> continuation there. Rows are placed in order, as TableMap fills them: each cell
/// takes the first column nothing above covers.
/// </summary>
internal sealed class TableCoverage
{
    /// <summary>Grid column → the covering cell's <c>colspan</c> and the rows it still covers.</summary>
    private Dictionary<int, (int Span, int Rows)> _above = [];

    /// <param name="firstColumn">Where the row's first cell sits: after its <c>w:gridBefore</c>, which the written row keeps.</param>
    public List<TableSlot> Place(List<Node> cells, int firstColumn)
    {
        var slots = new List<TableSlot>();
        var below = new Dictionary<int, (int Span, int Rows)>();
        var column = firstColumn;
        var index = 0;

        while (index < cells.Count || _above.Keys.Any(start => start >= column))
        {
            if (_above.TryGetValue(column, out var covering))
            {
                slots.Add(new TableSlot(null, covering.Span));
                if (covering.Rows > 1) below[column] = (covering.Span, covering.Rows - 1);
                column += covering.Span;
                continue;
            }

            if (index >= cells.Count)
            {
                column++;
                continue;
            }

            var cell = cells[index++];
            var span = SpanOf(cell);
            if (RowSpanOf(cell) > 1) below[column] = (span, RowSpanOf(cell) - 1);
            slots.Add(new TableSlot(cell, span));
            column += span;
        }

        _above = below;
        return slots;
    }

    public static int SpanOf(Node cell) => Attr.Int(cell, "colspan") is > 1 and var span ? span : 1;

    public static int RowSpanOf(Node cell) => Attr.Int(cell, "rowspan") is > 1 and var span ? span : 1;
}

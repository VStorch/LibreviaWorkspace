using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <c>w:tblGrid</c>, <c>w:tblW</c>, <c>w:tblLayout</c> and each cell's <c>w:tcW</c>. **Only
/// rewrites on change**: twip→pixel rounds, and a column whose pixel width did not change goes back
/// with its original twips. <c>w:tblW</c> in twips with fixed layout makes Word draw the requested
/// width, like the screen's <c>table-layout: fixed</c>.
internal sealed class TableGridWriter(Inventory inventory, int usableWidthPx)
{
    /// <summary>
    /// A little over twelve millimetres; below that, a new column gets the average of the others.
    /// </summary>
    private const int MinSharePx = 48;

    internal const string PartialGridLoss =
        "largura de coluna de uma tabela com linhas que não ocupam todas as colunas";

    public void Apply(Table table, List<Node> rows, Table? original)
    {
        var columns = Columns(rows);
        if (columns == 0) return;

        var declared = DeclaredWidths(rows, columns);
        var grid = original is null ? [] : GridTwips(original);

        // The first row does not cover the grid (`w:gridBefore`, `w:gridAfter`): the file's grid
        // stays, and the width change goes to the warning.
        if (original is not null && grid.Count > 0 && IsPartial(original, grid.Count))
        {
            var offset = GridBeforeOf(original);
            var same = offset + columns <= grid.Count
                       && declared.Select((width, index) => width is null || width == TableLook.ToPixels(grid[offset + index]))
                           .All(equal => equal);
            if (!same) inventory.NoteLoss(PartialGridLoss);
            return;
        }

        if (declared.All(width => width is null))
        {
            if (grid.Count == columns) return;

            // A freshly inserted table: equal columns across the text column, as in Word.
            var total = grid.Count > 0 ? (int)grid.Sum() : TableLook.ToTwips(usableWidthPx);
            WriteGrid(table, [.. Enumerable.Repeat((long)Math.Max(1, total / columns), columns)]);
            return;
        }

        var twips = Matched(Completed(declared, grid), grid);
        if (grid.SequenceEqual(twips)) return;

        WriteGrid(table, twips);
        WriteCellWidths(table, rows, twips);
    }

    /// <summary>
    /// TableKit only puts <c>colwidth</c> on the dragged column, and an inserted one arrives with
    /// <c>0</c> or <c>null</c>. The math is the screen's: an unmeasured column keeps the file's, or
    /// shares what is left.
    /// </summary>
    private List<int> Completed(List<int?> declared, List<long> grid)
    {
        if (grid.Count == declared.Count)
        {
            return [.. declared.Select((width, index) => width ?? TableLook.ToPixels(grid[index]))];
        }

        var known = declared.Where(width => width is not null).Select(width => width!.Value).ToList();
        var unknown = declared.Count - known.Count;
        if (unknown == 0) return known;

        var total = grid.Count > 0 ? grid.Sum(TableLook.ToPixels) : usableWidthPx;
        var share = (total - known.Sum()) / unknown;
        if (share < MinSharePx) share = Math.Max(MinSharePx, (int)Math.Round(known.Average()));

        return [.. declared.Select(width => width ?? share)];
    }

    /// <summary>
    /// By index with the same column count; otherwise in order: an inserted one is a new measure
    /// and a deleted one is skipped.
    /// </summary>
    private static List<long> Matched(List<int> widths, List<long> grid)
    {
        if (widths.Count == grid.Count)
        {
            return [.. widths.Select((width, index) =>
                TableLook.ToPixels(grid[index]) == width ? grid[index] : TableLook.ToTwips(width))];
        }

        var twips = new List<long>(widths.Count);
        var next = 0;

        foreach (var width in widths)
        {
            var found = next;
            while (widths.Count < grid.Count && found < grid.Count && TableLook.ToPixels(grid[found]) != width) found++;

            if (found < grid.Count && TableLook.ToPixels(grid[found]) == width)
            {
                twips.Add(grid[found]);
                next = found + 1;
                continue;
            }

            twips.Add(TableLook.ToTwips(width));
        }

        return twips;
    }

    private static void WriteGrid(Table table, List<long> twips)
    {
        var grid = new TableGrid(twips.Select(width => new GridColumn
        {
            Width = width.ToString(CultureInfo.InvariantCulture),
        }));

        var properties = table.GetFirstChild<TableProperties>();
        if (properties is null)
        {
            properties = new TableProperties();
            table.InsertAt(properties, 0);
        }

        properties.TableWidth = new TableWidth
        {
            Width = twips.Sum().ToString(CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa,
        };
        properties.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };

        if (table.GetFirstChild<TableGrid>() is { } existing) existing.Remove();
        table.InsertAfter(grid, properties);
    }

    /// <summary>
    /// An old <c>w:tcW</c> would contradict the new grid. Only on rows covering the whole grid.
    /// </summary>
    private static void WriteCellWidths(Table table, List<Node> rows, List<long> twips)
    {
        var built = table.Elements<TableRow>().ToList();
        for (var index = 0; index < built.Count && index < rows.Count; index++)
        {
            var cells = built[index].Elements<TableCell>().ToList();
            var modelled = rows[index].Content ?? [];
            if (cells.Count != modelled.Count || modelled.Sum(SpanOf) != twips.Count) continue;

            var column = 0;
            for (var position = 0; position < cells.Count; position++)
            {
                var span = SpanOf(modelled[position]);
                var width = twips.Skip(column).Take(span).Sum();
                column += span;

                var properties = cells[position].TableCellProperties ??= new TableCellProperties();
                properties.TableCellWidth = new TableCellWidth
                {
                    Width = width.ToString(CultureInfo.InvariantCulture),
                    Type = TableWidthUnitValues.Dxa,
                };
            }
        }
    }

    /// <summary>
    /// From the first row, as TableKit reads it for <c>colgroup</c>; <c>null</c> for an unmeasured
    /// column.
    /// </summary>
    private static List<int?> DeclaredWidths(List<Node> rows, int columns)
    {
        var widths = new List<int?>(columns);

        foreach (var cell in rows.FirstOrDefault()?.Content ?? [])
        {
            var span = SpanOf(cell);
            var declared = Attr.Node(cell, "colwidth") as JsonArray;

            for (var index = 0; index < span; index++)
            {
                var width = declared is not null && index < declared.Count ? declared[index] : null;
                var value = width?.GetValueKind() == JsonValueKind.Number ? width.GetValue<int>() : 0;
                widths.Add(value > 0 ? value : null);
            }
        }

        return widths;
    }

    private static int SpanOf(Node cell) => Attr.Int(cell, "colspan") is > 1 and var span ? span : 1;

    private static int Columns(List<Node> rows) => (rows.FirstOrDefault()?.Content ?? []).Sum(SpanOf);

    /// <summary>Empty if some measure cannot be read.</summary>
    private static List<long> GridTwips(Table original)
    {
        var widths = new List<long>();
        foreach (var column in original.GetFirstChild<TableGrid>()?.Elements<GridColumn>() ?? [])
        {
            if (!long.TryParse(column.Width?.Value, out var twips) || twips <= 0) return [];
            widths.Add(twips);
        }

        return widths;
    }

    /// <summary>
    /// The model does not represent <c>w:gridBefore</c>, <c>w:gridAfter</c> or rows with fewer
    /// cells.
    /// </summary>
    private static bool IsPartial(Table original, int gridColumns)
    {
        var row = original.Elements<TableRow>().FirstOrDefault();
        if (row is null) return false;

        var after = row.TableRowProperties?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0;
        var spans = row.Elements<TableCell>()
            .Sum(cell => cell.TableCellProperties?.GridSpan?.Val?.Value is > 1 and var span ? span : 1);
        return GridBeforeOf(original) > 0 || after > 0 || spans != gridColumns;
    }

    private static int GridBeforeOf(Table original) =>
        original.Elements<TableRow>().FirstOrDefault()?.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value
            is > 0 and var before
            ? before
            : 0;
}

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// <c>w:tblGrid</c>, <c>w:tblW</c>, <c>w:tblLayout</c> e o <c>w:tcW</c> de cada célula. **Só
/// reescreve quando muda**: twip→pixel arredonda, e a coluna cujo pixel não mudou
/// volta com o twip original. <c>w:tblW</c> em twips com layout fixo faz o Word
/// desenhar a largura pedida, como o <c>table-layout: fixed</c> da tela.
/// </summary>
internal sealed class TableGridWriter(Inventory inventory, int usableWidthPx)
{
    /// <summary>Doze milímetros e pouco; abaixo disso, a coluna nova ganha a média das outras.</summary>
    private const int MinSharePx = 48;

    internal const string PartialGridLoss =
        "largura de coluna de uma tabela com linhas que não ocupam todas as colunas";

    public void Apply(Table table, List<Node> rows, Table? original)
    {
        var columns = Columns(rows);
        if (columns == 0) return;

        var declared = DeclaredWidths(rows, columns);
        var grid = original is null ? [] : GridTwips(original);

        // A primeira linha não cobre a grade (`w:gridBefore`, `w:gridAfter`): a do
        // arquivo fica, e a mudança de largura vai ao aviso.
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

            // Tabela recém-inserida: colunas iguais na coluna de texto, como o Word.
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
    /// O TableKit põe <c>colwidth</c> só na coluna arrastada, e a inserida chega com
    /// <c>0</c> ou <c>null</c>. A conta é a da tela: a sem medida fica com a do arquivo, ou
    /// divide o que sobra.
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
    /// Pelo índice com o mesmo número de colunas; com outro, na ordem: a inserida é
    /// medida nova e a apagada é pulada.
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

    /// <summary>O <c>w:tcW</c> antigo contradiria a grade nova. Só nas linhas que cobrem a grade inteira.</summary>
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

    /// <summary>Da primeira linha, como o TableKit lê para o <c>colgroup</c>; <c>null</c> na coluna sem medida.</summary>
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

    /// <summary>Vazia se alguma medida não se lê.</summary>
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

    /// <summary>O modelo não representa <c>w:gridBefore</c>, <c>w:gridAfter</c> nem linha com menos células.</summary>
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

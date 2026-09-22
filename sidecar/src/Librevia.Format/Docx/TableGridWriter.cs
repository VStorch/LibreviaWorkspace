using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A grade de colunas da tabela gravada — `w:tblGrid`, `w:tblW`, `w:tblLayout` e
/// o `w:tcW` de cada célula.
/// </summary>
/// <remarks>
/// Saiu de <see cref="TableWriter"/> porque é a parte da tabela que faz conta:
/// completar a largura que o editor não declarou, casar pixel com twip e decidir
/// se a grade do arquivo ainda serve. Misturada à cópia por posição de linhas e
/// células, as duas lógicas se escondiam uma na outra.
///
/// O editor já deixava arrastar a divisória das colunas, e o número novo morria
/// aqui: o `w:tblGrid` voltava do arquivo por posição e ninguém o comparava com o
/// modelo. Era perda silenciosa.
///
/// **Só reescreve quando muda.** A conversão twip→pixel arredonda — 2000 twips
/// são 133,33 px —, então regravar a grade a cada salvamento mexeria na medida
/// de uma tabela que ninguém tocou. Igual ao que o arquivo tem, o XML original
/// volta intacto; e a coluna cujo pixel não mudou volta com o twip original.
///
/// Com a grade vem `w:tblW` em twips e `w:tblLayout` fixo: é o par que faz o Word
/// desenhar a largura pedida em vez de redistribuir as colunas pelo conteúdo —
/// que é o que a tela faz, com `table-layout: fixed`.
/// </remarks>
internal sealed class TableGridWriter(Inventory inventory, int usableWidthPx)
{
    /// <summary>
    /// A coluna mais estreita que a conta de "o que sobra" pode dar à coluna sem
    /// medida: doze milímetros e pouco. Abaixo disso — a tabela já ocupava a
    /// coluna de texto inteira e alguém inseriu mais uma — a coluna nova ganha a
    /// média das outras, que é mais perto do que o Word faz do que um fio de um
    /// pixel.
    /// </summary>
    private const int MinSharePx = 48;

    /// <summary>Largura de coluna mudada numa tabela cuja primeira linha não cobre a grade.</summary>
    internal const string PartialGridLoss =
        "largura de coluna de uma tabela com linhas que não ocupam todas as colunas";

    public void Apply(Table table, List<Node> rows, Table? original)
    {
        var columns = Columns(rows);
        if (columns == 0) return;

        var declared = DeclaredWidths(rows, columns);
        var grid = original is null ? [] : GridTwips(original);

        // A primeira linha do arquivo não cobre a grade toda — `w:gridBefore`,
        // `w:gridAfter` —, e a grade que o modelo descreve é só um pedaço da do
        // arquivo. Regravá-la com menos colunas mudava a forma da tabela na
        // primeira correção de texto. A grade do arquivo fica; se a largura mudou
        // de fato, a mudança não chega, e o aviso diz isso.
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
            // Nada declarado e a grade do arquivo ainda cabe: é ela que vale.
            if (grid.Count == columns) return;

            // Tabela recém-inserida: o editor não declara largura nenhuma, e o
            // esquema quer a grade. Colunas iguais na coluna de texto é o que o
            // Word faz ao inserir.
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
    /// A largura das colunas que o editor não declarou, completada.
    /// </summary>
    /// <remarks>
    /// O TableKit põe `colwidth` só na coluna arrastada, e depois de inserir uma
    /// coluna a nova chega com `0` ou `null`. Antes daqui a grade parcial era
    /// descartada inteira, e a largura que a pessoa arrastou não chegava ao
    /// arquivo. A conta é a da tela: com a grade do arquivo no mesmo formato, a
    /// coluna sem medida fica com a dela; sem isso, as sem medida dividem o que
    /// sobra — da tabela do arquivo, ou da coluna de texto numa tabela nova.
    /// </remarks>
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
    /// Pixel → twip, devolvendo o twip original a cada coluna cujo pixel não mudou.
    /// </summary>
    /// <remarks>
    /// Com o mesmo número de colunas, pelo índice. Com outro — coluna inserida ou
    /// apagada —, o índice já não aponta para a mesma coluna do arquivo, e a conta
    /// anda na ordem: a coluna que casa com a próxima original leva o twip dela; a
    /// inserida é medida nova e não consome original nenhuma; a apagada é pulada.
    /// </remarks>
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
    /// A largura declarada em cada célula, que é o outro lugar de onde o Word
    /// tira a medida: deixada com o número antigo, ela contradiz a grade nova e o
    /// Word escolhe uma das duas sem avisar.
    /// </summary>
    /// <remarks>
    /// Só nas linhas que cobrem a grade inteira: numa linha de outro formato não
    /// há como saber a que colunas cada célula corresponde.
    /// </remarks>
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
    /// A largura que o modelo declara para cada coluna da grade, lida da primeira
    /// linha; <c>null</c> na coluna sem medida.
    /// </summary>
    /// <remarks>
    /// Da primeira linha porque é de lá que o editor também a lê para desenhar o
    /// `colgroup` — ver `updateColumns` do TableKit.
    /// </remarks>
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

    /// <summary>Quantas colunas de grade a tabela tem, somando os `colspan`.</summary>
    private static int Columns(List<Node> rows) => (rows.FirstOrDefault()?.Content ?? []).Sum(SpanOf);

    /// <summary>A grade do arquivo em twips, como está escrita; vazia se alguma medida não se lê.</summary>
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
    /// A primeira linha do arquivo deixa colunas da grade de fora — pula algumas
    /// no começo (`w:gridBefore`), deixa outras no fim (`w:gridAfter`) ou tem
    /// menos células do que a grade? O modelo não representa nenhum dos três.
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

    /// <summary>Quantas colunas a primeira linha do arquivo pula antes da primeira célula.</summary>
    private static int GridBeforeOf(Table original) =>
        original.Elements<TableRow>().FirstOrDefault()?.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value
            is > 0 and var before
            ? before
            : 0;
}

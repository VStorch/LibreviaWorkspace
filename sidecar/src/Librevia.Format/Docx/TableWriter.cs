using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Tabela → <c>w:tbl</c>, **por posição**: cada <c>w:tblPr</c>, <c>w:tblGrid</c>,
/// <c>w:trPr</c> e <c>w:tcPr</c> do arquivo volta ao lugar, e só os parágrafos são
/// regravados. Linha ou coluna no fim ganha estrutura nova; inserir no meio ou
/// remover desfaz a correspondência, e isso vai ao inventário.
/// </summary>
internal sealed class TableWriter(
    Inventory inventory,
    Func<Node, OpenXmlElement?, IEnumerable<OpenXmlElement>> writeBlock,
    int usableWidthPx,
    bool revisions = true)
{
    /// <summary>
    /// Removida, a linha leva largura, mesclagem e sombreamento; inserida no meio,
    /// desloca a estrutura, e o <c>w:vMerge</c> fora de lugar faz o Word acusar tabela
    /// corrompida.
    /// </summary>
    private const string ShiftedStructure =
        "largura, mesclagem ou sombreamento de parte de uma tabela que você editou";

    /// <summary>O gravador ainda não a escreve.</summary>
    internal const string VerticalMergeLoss = "mesclagem vertical de células feita no editor";

    private readonly TableGridWriter _grid = new(inventory, usableWidthPx);

    public Table Write(Node node, Table? original)
    {
        var table = new Table();

        // Antes da primeira linha: propriedades, grade e a revisão da tabela.
        var interleaved = Interleaved<TableRow>(original);
        foreach (var setting in Strays(interleaved, -1)) table.AppendChild(setting.CloneNode(true));
        if (original is null) table.AppendChild(DefaultProperties());

        var originalRows = original?.Elements<TableRow>().ToList() ?? [];
        var rows = node.Content ?? [];

        // O modelo não carrega identidade de linha: com contagens diferentes, não há como saber onde mudou.
        var aligned = original is null || originalRows.Count == rows.Count;
        if (!aligned) inventory.NoteLoss(ShiftedStructure);

        for (var index = 0; index < rows.Count; index++)
        {
            table.AppendChild(WriteRow(rows[index], originalRows.ElementAtOrDefault(index), aligned));
            foreach (var between in Strays(interleaved, index)) table.AppendChild(between.CloneNode(true));
        }

        // O que vinha depois de uma linha que não existe mais fecha a tabela.
        for (var index = rows.Count; index < originalRows.Count; index++)
        {
            foreach (var orphan in Strays(interleaved, index)) table.AppendChild(orphan.CloneNode(true));
        }

        _grid.Apply(table, rows, original);
        return table;
    }

    /// <summary>
    /// Os filhos que não são <typeparamref name="T"/>, pela posição: <c>-1</c> antes do
    /// primeiro, <c>n</c> depois do de índice <c>n</c>. Entre as linhas moram marcadores,
    /// <c>w:sdt</c> e revisões, que não podem ir todos para o começo.
    /// </summary>
    private static Dictionary<int, List<OpenXmlElement>> Interleaved<T>(OpenXmlElement? parent)
        where T : OpenXmlElement
    {
        var map = new Dictionary<int, List<OpenXmlElement>>();
        if (parent is null) return map;

        var position = -1;
        foreach (var child in parent.ChildElements)
        {
            if (child is T)
            {
                position++;
                continue;
            }

            if (!map.TryGetValue(position, out var bucket)) map[position] = bucket = [];
            bucket.Add(child);
        }

        return map;
    }

    private static List<OpenXmlElement> Strays(Dictionary<int, List<OpenXmlElement>> map, int position) =>
        map.TryGetValue(position, out var found) ? found : [];

    /// <summary>A tabela criada aqui; a largura e o <c>w:tblLayout</c> vêm com a grade (<see cref="TableGridWriter"/>).</summary>
    private static TableProperties DefaultProperties() => new(
        // Na ordem do esquema, senão o Word recusa o documento.
        new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));

    /// <param name="tableAligned">Falso: a estrutura desta linha pode ser de outra.</param>
    private TableRow WriteRow(Node rowNode, TableRow? original, bool tableAligned)
    {
        var row = new TableRow();

        // `w:trPr` traz a altura e o `w:tblHeader`; `w:tblPrEx`, as exceções da linha.
        var interleaved = Interleaved<TableCell>(original);
        foreach (var setting in Strays(interleaved, -1)) row.AppendChild(setting.CloneNode(true));

        var originalCells = original?.Elements<TableCell>().ToList() ?? [];
        var cells = rowNode.Content ?? [];

        var aligned = original is null || originalCells.Count == cells.Count;
        if (!aligned) inventory.NoteLoss(ShiftedStructure);

        for (var index = 0; index < cells.Count; index++)
        {
            row.AppendChild(WriteCell(
                cells[index],
                originalCells.ElementAtOrDefault(index),
                aligned && tableAligned));
            foreach (var between in Strays(interleaved, index)) row.AppendChild(between.CloneNode(true));
        }

        for (var index = cells.Count; index < originalCells.Count; index++)
        {
            foreach (var orphan in Strays(interleaved, index)) row.AppendChild(orphan.CloneNode(true));
        }

        // Linha toda de `tableHeader` é o `w:tblHeader`.
        ApplyHeader(row, cells.Count > 0 && cells.All(cell => cell.Type == "tableHeader"));
        if (revisions) ApplyRowRevision(row, Attr.Node(rowNode, "rowRevision"));
        return row;
    }

    /// <summary><c>w:ins</c>/<c>w:del</c> no fim do <c>w:trPr</c>, antes só do <c>w:trPrChange</c>.</summary>
    private static void ApplyRowRevision(TableRow row, System.Text.Json.Nodes.JsonNode? wanted)
    {
        var properties = row.TableRowProperties;
        if (properties is null)
        {
            if (wanted is null) return;
            properties = new TableRowProperties();
            row.PrependChild(properties);
        }

        Revisions.ApplyBlock(properties, wanted, created =>
        {
            var change = properties.GetFirstChild<TableRowPropertiesChange>();
            if (change is null) properties.AppendChild(created);
            else properties.InsertBefore(created, change);
        });

        if (!properties.HasChildren) properties.Remove();
    }

    /// <summary>Só quando o estado muda: o <c>w:trPr</c> traz também a altura e a revisão.</summary>
    private static void ApplyHeader(TableRow row, bool wanted)
    {
        var properties = row.TableRowProperties;
        var current = properties?.GetFirstChild<TableHeader>();
        // Presente sem `w:val` já é "sim".
        var declared = current is not null && !IsOff(current);
        if (declared == wanted) return;

        current?.Remove();
        if (!wanted) return;

        if (properties is null)
        {
            properties = new TableRowProperties();
            row.TableRowProperties = properties;
        }

        // Antes da revisão, que fecha o `w:trPr` no esquema.
        var revision = properties.ChildElements.FirstOrDefault(child =>
            child is Inserted or Deleted or TableRowPropertiesChange);
        if (revision is null) properties.AppendChild(new TableHeader());
        else properties.InsertBefore(new TableHeader(), revision);
    }

    private static bool IsOff(TableHeader header) =>
        header.Val is { } value && value.Value == OnOffOnlyValues.Off;

    private static List<OpenXmlElement> ChildrenOf(OpenXmlElement? element) =>
        element is null ? [] : [.. element.ChildElements];

    /// <param name="aligned">Se esta é de fato a célula do arquivo nesta posição.</param>
    private TableCell WriteCell(Node cellNode, TableCell? original, bool aligned)
    {
        var cell = new TableCell();

        var properties = original?.TableCellProperties?.CloneNode(true) as TableCellProperties
                         ?? new TableCellProperties();
        var span = Attr.Int(cellNode, "colspan");

        // Só o que o modelo representa (`colspan`, sombreamento, bordas): `w:vMerge`,
        // margem interna e alinhamento vertical ficam.
        if (span is > 1) properties.GridSpan = new GridSpan { Val = span };
        else if (properties.GridSpan is not null) properties.GridSpan = null;

        // Mesclagem vertical deslocada faz o Word acusar tabela corrompida: sai, e a perda já foi declarada.
        if (!aligned) properties.RemoveAllChildren<VerticalMerge>();

        // A mesclagem vertical feita na tela vira `rowspan`, que este gravador não escreve: vai ao inventário.
        if (Attr.Int(cellNode, "rowspan") is > 1) inventory.NoteLoss(VerticalMergeLoss);

        // Só quando diferem do arquivo — ver TableLook.
        TableLook.ApplyShading(properties, Attr.String(cellNode, "shading"), inventory);
        TableLook.ApplyBorders(properties, Attr.String(cellNode, "borders"), inventory);

        if (properties.HasChildren) cell.TableCellProperties = properties;

        var children = ChildrenOf(original);
        var blocks = children.Where(child => child is Paragraph or Table).ToList();

        // Um controle de conteúdo ou um campo entre os parágrafos voltaria como nada.
        if (children.Any(child => child is not (Paragraph or Table or TableCellProperties)))
        {
            inventory.NoteLoss("conteúdo especial de uma célula que você editou");
        }

        var position = 0;
        foreach (var child in cellNode.Content ?? [])
        {
            var source = blocks.ElementAtOrDefault(position);
            position++;

            foreach (var element in writeBlock(child, source)) cell.AppendChild(element);
        }

        // O `w:tc` tem de **terminar** em `w:p`, e a tabela aninhada pode terminar em `w:tbl`.
        if (cell.LastChild is not Paragraph) cell.AppendChild(new Paragraph());
        return cell;
    }
}

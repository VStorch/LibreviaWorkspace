using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Tabela do editor → `w:tbl`, preservando o que o modelo não representa.
/// </summary>
/// <remarks>
/// O modelo do editor sabe muito pouco sobre tabela: linhas, células e quantas
/// colunas cada célula ocupa. O arquivo sabe muito mais — largura de cada coluna,
/// célula mesclada na vertical, sombreamento, estilo de tabela, bordas de cada
/// lado e a linha que se repete no alto de cada página.
///
/// Escrever a tabela a partir do modelo apagava tudo isso e punha bordas finas
/// iguais em toda parte: corrigir uma palavra dentro de uma célula desmontava a
/// tabela inteira. A saída é a mesma da gravação cirúrgica: **por posição**, cada
/// `w:tblPr`, `w:tblGrid`, `w:trPr` e `w:tcPr` do arquivo volta para o lugar de
/// onde veio, e só o conteúdo dos parágrafos é regravado.
///
/// Quando a pessoa acrescenta linha ou coluna no fim não há original para aquela
/// posição, e aí a estrutura nova é gerada. Quando ela insere no meio ou
/// **remove**, a correspondência por posição deixa de valer para o resto da
/// tabela: o que havia no arquivo se perde ou sai no lugar errado. Isso entra no
/// inventário, porque é o único jeito de a perda não ser silenciosa.
/// </remarks>
internal sealed class TableWriter(
    Inventory inventory,
    Func<Node, OpenXmlElement?, IEnumerable<OpenXmlElement>> writeBlock,
    int usableWidthPx)
{
    /// <summary>
    /// O aviso de que a correspondência por posição deixou de valer.
    /// </summary>
    /// <remarks>
    /// Linha ou célula **removida** leva embora a largura, a mesclagem e o
    /// sombreamento que estavam no arquivo. Linha ou célula **inserida no meio**
    /// é igualmente grave, e antes daqui saía calada: a estrutura de uma linha
    /// passa a valer para a de baixo, e o `w:vMerge` fora de lugar é justamente o
    /// caso em que o Word acusa tabela corrompida.
    /// </remarks>
    private const string ShiftedStructure =
        "largura, mesclagem ou sombreamento de parte de uma tabela que você editou";

    /// <summary>A mesclagem vertical feita na tela, que o gravador ainda não escreve.</summary>
    internal const string VerticalMergeLoss = "mesclagem vertical de células feita no editor";

    private readonly TableGridWriter _grid = new(inventory, usableWidthPx);

    public Table Write(Node node, Table? original)
    {
        var table = new Table();

        // Tudo o que vem antes da primeira linha é configuração da tabela:
        // propriedades, grade de colunas e, em documento do Word, também a
        // marcação de revisão da própria tabela.
        var interleaved = Interleaved<TableRow>(original);
        foreach (var setting in Strays(interleaved, -1)) table.AppendChild(setting.CloneNode(true));
        if (original is null) table.AppendChild(DefaultProperties());

        var originalRows = original?.Elements<TableRow>().ToList() ?? [];
        var rows = node.Content ?? [];

        // Contagens diferentes querem dizer que a linha de índice `n` do modelo
        // pode não ser a de índice `n` do arquivo, e a gravação por posição
        // deixa de ser fiel. Não há como saber **onde** a pessoa inseriu ou
        // removeu: o modelo do editor não carrega identidade de linha.
        var aligned = original is null || originalRows.Count == rows.Count;
        if (!aligned) inventory.NoteLoss(ShiftedStructure);

        for (var index = 0; index < rows.Count; index++)
        {
            table.AppendChild(WriteRow(rows[index], originalRows.ElementAtOrDefault(index), aligned));
            foreach (var between in Strays(interleaved, index)) table.AppendChild(between.CloneNode(true));
        }

        // O que vinha depois de uma linha que não existe mais fecha a tabela: é
        // onde ele menos desloca o que abraçava.
        for (var index = rows.Count; index < originalRows.Count; index++)
        {
            foreach (var orphan in Strays(interleaved, index)) table.AppendChild(orphan.CloneNode(true));
        }

        _grid.Apply(table, rows, original);
        return table;
    }

    /// <summary>
    /// Os filhos que não são <typeparamref name="T"/>, agrupados pela posição em
    /// que estavam: <c>-1</c> é o que vem antes do primeiro, e <c>n</c> o que
    /// vinha depois do de índice <c>n</c>.
    /// </summary>
    /// <remarks>
    /// `w:tblPr` e `w:tblGrid` abrem a tabela, e `w:trPr` abre a linha — esses são
    /// os de <c>-1</c>. Mas entre as linhas moram outras coisas: um
    /// `w:bookmarkStart`, um `w:sdt`, a marcação de revisão da tabela. Levar todos
    /// para antes da primeira linha — o que este escritor fazia — encurta o
    /// marcador até o vazio e tira o controle de conteúdo de onde ele valia.
    /// </remarks>
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

    /// <summary>Bordas visíveis para a tabela criada aqui dentro, que não tem original.</summary>
    /// <remarks>
    /// A largura e o `w:tblLayout` entram em <see cref="TableGridWriter"/>, junto com a
    /// grade: sem grade não há largura a declarar.
    /// </remarks>
    private static TableProperties DefaultProperties() => new(
        // Na ordem da sequência do esquema — topo, esquerda, baixo, direita, e só
        // então as de dentro. Fora dela o Word recusa o documento, e nenhum teste
        // pegava isso porque nenhum gravava uma tabela criada na tela.
        new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));

    /// <param name="tableAligned">
    /// Se as linhas do modelo ainda correspondem às do arquivo. Falso arrasta
    /// para as células: a estrutura desta linha pode ser de outra.
    /// </param>
    private TableRow WriteRow(Node rowNode, TableRow? original, bool tableAligned)
    {
        var row = new TableRow();

        // `w:trPr` traz a altura da linha e o `w:tblHeader` que a repete no alto
        // de cada página; `w:tblPrEx` são as exceções de formatação da linha.
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

        // Linha toda de `tableHeader` é a linha de cabeçalho do editor, e no
        // arquivo isso é o `w:tblHeader` do `w:trPr`: a linha se repete no alto de
        // cada página. Sem esta escrita o botão da tela não chegava ao arquivo.
        ApplyHeader(row, cells.Count > 0 && cells.All(cell => cell.Type == "tableHeader"));
        return row;
    }

    /// <summary>Liga ou desliga o `w:tblHeader` da linha, sem mexer no resto do `w:trPr`.</summary>
    /// <remarks>
    /// Só escreve quando o estado muda. O `w:trPr` do arquivo traz também a altura
    /// da linha e a marcação de revisão, e recriá-lo para ligar uma bandeira
    /// levaria os dois embora.
    /// </remarks>
    private static void ApplyHeader(TableRow row, bool wanted)
    {
        var properties = row.TableRowProperties;
        var current = properties?.GetFirstChild<TableHeader>();
        // `w:tblHeader` presente sem `w:val` já é "sim"; só `w:val="false"` nega.
        var declared = current is not null && !IsOff(current);
        if (declared == wanted) return;

        current?.Remove();
        if (!wanted) return;

        if (properties is null)
        {
            properties = new TableRowProperties();
            row.TableRowProperties = properties;
        }

        // Antes da marcação de revisão, que fecha o `w:trPr`: as propriedades da
        // linha vêm primeiro e `w:ins`, `w:del` e `w:trPrChange` por último.
        // Anexado depois delas, o documento saía fora do esquema.
        var revision = properties.ChildElements.FirstOrDefault(child =>
            child is Inserted or Deleted or TableRowPropertiesChange);
        if (revision is null) properties.AppendChild(new TableHeader());
        else properties.InsertBefore(new TableHeader(), revision);
    }

    private static bool IsOff(TableHeader header) =>
        header.Val is { } value && value.Value == OnOffOnlyValues.Off;

    private static List<OpenXmlElement> ChildrenOf(OpenXmlElement? element) =>
        element is null ? [] : [.. element.ChildElements];

    /// <param name="aligned">
    /// Se esta célula é de fato a célula do arquivo que está nesta posição.
    /// </param>
    private TableCell WriteCell(Node cellNode, TableCell? original, bool aligned)
    {
        var cell = new TableCell();

        var properties = original?.TableCellProperties?.CloneNode(true) as TableCellProperties
                         ?? new TableCellProperties();
        var span = Attr.Int(cellNode, "colspan");

        // O modelo sabe do `colspan`, do sombreamento e das bordas, e os três
        // podem ter mudado na tela — o `rowspan` dele é sempre 1, porque a
        // mesclagem vertical mora no arquivo. Sobrepor só o que o modelo
        // representa mantém intactos o `w:vMerge`, a margem interna e o alinhamento
        // vertical que vieram do arquivo.
        if (span is > 1) properties.GridSpan = new GridSpan { Val = span };
        else if (properties.GridSpan is not null) properties.GridSpan = null;

        // Mesclagem vertical deslocada é o caso em que o Word acusa tabela
        // corrompida — o `w:vMerge w:val="continue"` de uma célula que já não
        // tem acima de si a que abriu a mesclagem. Quando as posições deixaram
        // de casar, ela é descartada: tabela sem mesclagem abre, tabela com
        // mesclagem errada não. A perda já está no inventário.
        if (!aligned) properties.RemoveAllChildren<VerticalMerge>();

        // Mesclagem vertical feita **na tela** — `mergeCells` sobre células de
        // linhas diferentes — vira `rowspan` no modelo, e este gravador não a
        // escreve como `w:vMerge`: a célula de baixo simplesmente não existe no
        // arquivo. Sai no inventário, e não em silêncio.
        if (Attr.Int(cellNode, "rowspan") is > 1) inventory.NoteLoss(VerticalMergeLoss);

        // As duas só são reescritas quando diferem do que o arquivo tem: é o que
        // permite ao `w:shd` com trama e ao `w:tcBorders` com estilo exótico
        // voltarem intactos na célula que ninguém formatou. Ver TableLook.
        TableLook.ApplyShading(properties, Attr.String(cellNode, "shading"), inventory);
        TableLook.ApplyBorders(properties, Attr.String(cellNode, "borders"), inventory);

        if (properties.HasChildren) cell.TableCellProperties = properties;

        var children = ChildrenOf(original);
        var blocks = children.Where(child => child is Paragraph or Table).ToList();

        // O que a célula tinha e o modelo não representa: um controle de conteúdo
        // ou um campo mora entre os parágrafos, e volta para o arquivo como nada.
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

        // O `w:tc` tem de **terminar** em `w:p`, e não só conter um: o Word
        // recusa o documento que não o faça. Desde que a leitura entrega a tabela
        // aninhada, a célula pode terminar em `w:tbl` — no schema do editor ela
        // aceita qualquer bloco —, e até aqui isso escapava só porque o original
        // trazia o parágrafo vazio do fim.
        if (cell.LastChild is not Paragraph) cell.AppendChild(new Paragraph());
        return cell;
    }
}

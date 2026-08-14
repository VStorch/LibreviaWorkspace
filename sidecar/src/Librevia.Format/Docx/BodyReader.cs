using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Librevia.Format.Docx;

/// <summary>Um bloco de primeiro nível do corpo, com identidade.</summary>
/// <param name="Oid">Id estável na ordem do documento: b1, b2, …</param>
/// <param name="Source">O elemento original, para gravar de volta sem tocar.</param>
/// <param name="Extracted">O que o editor vê.</param>
public sealed record Block(string Oid, OpenXmlElement Source, Node Extracted);

/// <summary>
/// Corpo do documento → nós do editor.
/// </summary>
/// <remarks>
/// Extração **best-effort e de mão única**: o que sai daqui alimenta a tela e o
/// PDF, nunca a gravação. Quem grava é <see cref="DocxWriter"/>, a partir do
/// XML original. Por isso um erro aqui é cosmético, não perda de dados — e é
/// isso que permite ser tolerante em vez de recusar o arquivo.
/// </remarks>
public sealed class BodyReader(MainDocumentPart part, Inventory inventory)
{
    /// <summary>Passo de recuo do Word: meia polegada.</summary>
    private const int TwipsPerIndentLevel = 720;

    private readonly NumberingReader _numbering = new(part);
    private int _nextId = 1;

    /// <summary>
    /// Percorre o corpo produzindo a árvore do editor e, em paralelo, a lista
    /// plana de blocos com identidade.
    /// </summary>
    /// <remarks>
    /// As duas saem juntas porque os blocos **apontam para nós de dentro da
    /// árvore**: um item de lista é um `w:p` no arquivo e um `listItem`
    /// aninhado na árvore. Montar a árvore primeiro e procurar os blocos depois
    /// exigiria adivinhar essa correspondência.
    /// </remarks>
    public (List<Node> Content, List<Block> Blocks) Read(Body body)
    {
        var content = new List<Node>();
        var blocks = new List<Block>();

        // Parágrafos numerados consecutivos viram uma lista só; a pilha guarda
        // as listas abertas, uma por nível de aninhamento.
        var openLists = new List<(Node List, string Kind, int Level)>();

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case Paragraph paragraph:
                {
                    var node = ReadParagraph(paragraph);
                    var list = node.Type == "pageBreak"
                        ? null
                        : _numbering.ListKindOf(paragraph.ParagraphProperties);

                    if (list is null)
                    {
                        openLists.Clear();
                        blocks.Add(NewBlock(element, node));
                        content.Add(node);
                        break;
                    }

                    var (kind, level) = (list.Value.Kind, list.Value.Level);
                    while (openLists.Count > 0 && openLists[^1].Level > level)
                    {
                        openLists.RemoveAt(openLists.Count - 1);
                    }

                    if (openLists.Count == 0 || openLists[^1].Level < level || openLists[^1].Kind != kind)
                    {
                        var listNode = Node.Of(kind);
                        listNode.Content = [];

                        if (openLists.Count > 0 && openLists[^1].Level < level)
                        {
                            // Lista aninhada mora dentro do último item da de fora.
                            var parentItems = openLists[^1].List.Content!;
                            (parentItems[^1].Content ??= []).Add(listNode);
                        }
                        else
                        {
                            openLists.Clear();
                            content.Add(listNode);
                        }

                        openLists.Add((listNode, kind, level));
                    }

                    // O id fica no `listItem`, não no parágrafo: é o item que
                    // corresponde a um `w:p` do arquivo.
                    var item = Node.Of("listItem");
                    item.Content = [node];
                    openLists[^1].List.Content!.Add(item);
                    blocks.Add(NewBlock(element, item));
                    break;
                }

                case Table table:
                {
                    openLists.Clear();
                    var node = ReadTable(table);
                    blocks.Add(NewBlock(element, node));
                    content.Add(node);
                    break;
                }

                case SectionProperties:
                    // Configuração de página: sai do fluxo e vira `page`.
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }

        // O ProseMirror recusa um documento sem nenhum bloco.
        if (content.Count == 0) content.Add(Node.Of("paragraph"));

        return (content, blocks);
    }

    private Block NewBlock(OpenXmlElement source, Node extracted)
    {
        var oid = "b" + _nextId++;
        extracted.With("oid", oid);
        return new Block(oid, source, extracted);
    }

    // --- parágrafos ---------------------------------------------------------

    private Node ReadParagraph(Paragraph paragraph)
    {
        var properties = paragraph.ParagraphProperties;
        var content = ReadInline(paragraph);

        // Uma quebra de página sozinha no parágrafo é o nó `pageBreak`, não um
        // parágrafo vazio com uma quebra dentro.
        if (content.Count == 1 && content[0].Type == "pageBreak")
        {
            return content[0];
        }

        var node = HeadingLevelOf(properties) is { } level
            ? Node.Of("heading").With("level", level)
            : Node.Of("paragraph");

        var alignment = AlignmentOf(properties);
        if (alignment is not null) node.With("textAlign", alignment);

        var indent = IndentOf(properties);
        if (indent > 0) node.With("indent", indent);

        node.Content = content.Count == 0 ? null : content;
        return node;
    }

    /// <summary>
    /// Nível do título a partir do estilo do parágrafo.
    /// </summary>
    /// <remarks>
    /// O corpus usa `Heading1`; o Word também grava `heading 1` e o LibreOffice
    /// `Ttulo1` (sem acento, porque o id do estilo não os aceita). Aceitar as
    /// três formas custa uma linha e evita que todo título vire texto normal.
    /// </remarks>
    private static int? HeadingLevelOf(ParagraphProperties? properties)
    {
        var style = properties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(style)) return null;

        var normalized = style.Replace(" ", string.Empty).ToLowerInvariant();
        foreach (var prefix in (string[])["heading", "ttulo", "titulo"])
        {
            if (normalized.StartsWith(prefix, StringComparison.Ordinal) &&
                int.TryParse(normalized[prefix.Length..], out var level) &&
                level is >= 1 and <= 6)
            {
                return level;
            }
        }

        return null;
    }

    private static string? AlignmentOf(ParagraphProperties? properties)
    {
        var value = properties?.Justification?.Val;
        if (value is null) return null;

        if (value == JustificationValues.Center) return "center";
        if (value == JustificationValues.Right) return "right";
        if (value == JustificationValues.Both || value == JustificationValues.Distribute) return "justify";
        return "left";
    }

    private static int IndentOf(ParagraphProperties? properties)
    {
        var left = properties?.Indentation?.Left?.Value;
        if (left is null || !int.TryParse(left, out var twips) || twips <= 0) return 0;
        return Math.Min(twips / TwipsPerIndentLevel, 10);
    }

    // --- conteúdo em linha --------------------------------------------------

    private List<Node> ReadInline(OpenXmlElement container, string? hyperlink = null)
    {
        var nodes = new List<Node>();

        foreach (var element in container.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    nodes.AddRange(ReadRun(run, hyperlink));
                    break;

                case Hyperlink link:
                    nodes.AddRange(ReadInline(link, HyperlinkTargetOf(link) ?? hyperlink));
                    break;

                case ParagraphProperties:
                case BookmarkStart:
                case BookmarkEnd:
                case ProofError:
                    break;

                // Comentários e revisões são preservados pelo XML original; o
                // editor não os mostra. Invisibilidade, não perda.
                case CommentRangeStart:
                case CommentRangeEnd:
                    inventory.NoteInvisible("comentários");
                    break;

                case InsertedRun inserted:
                    inventory.NoteInvisible("controle de alterações");
                    nodes.AddRange(ReadInline(inserted, hyperlink));
                    break;

                case DeletedRun:
                    // Texto marcado como excluído não deve aparecer na tela.
                    inventory.NoteInvisible("controle de alterações");
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }

        return nodes;
    }

    private IEnumerable<Node> ReadRun(Run run, string? hyperlink)
    {
        var marks = RunReader.MarksOf(run.RunProperties, hyperlink);

        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case Text text:
                    if (text.Text.Length > 0)
                    {
                        yield return new Node { Type = "text", Text = text.Text, Marks = marks };
                    }

                    break;

                case TabChar:
                    yield return new Node { Type = "text", Text = "\t", Marks = marks };
                    break;

                case Break br:
                    yield return br.Type is not null && br.Type.Value == BreakValues.Page
                        ? Node.Of("pageBreak")
                        : Node.Of("hardBreak");
                    break;

                case DocumentFormat.OpenXml.Wordprocessing.Drawing:
                    if (ReadImage(element) is { } image) yield return image;
                    break;

                case Picture picture:
                    // VML antigo. No corpus só aparece nos cabeçalhos, mas um
                    // documento do Word 2003 traz imagens assim no corpo.
                    if (ReadImage(picture) is { } legacy) yield return legacy;
                    break;

                case RunProperties:
                case LastRenderedPageBreak:
                    break;

                case FieldChar:
                case FieldCode:
                    // `PAGE` e afins: o valor só existe na paginação. O texto
                    // que o Word deixou em cache vem no run seguinte.
                    inventory.NoteInvisible("campos calculados (como sumário e número de página)");
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }
    }

    // --- imagens ------------------------------------------------------------

    /// <summary>
    /// Imagem → nó com data URI.
    /// </summary>
    /// <remarks>
    /// Toda imagem do corpus é `wp:anchor` centralizada com largura igual à do
    /// texto — o jeito do LibreOffice gravar "imagem no próprio parágrafo".
    /// Tratá-la como flutuante produziria layout pior, não melhor. Ver
    /// docs/01-corpus-docx.md, Descoberta 3.
    /// </remarks>
    private Node? ReadImage(OpenXmlElement drawing)
    {
        var blip = drawing.Descendants<Drawing.Blip>().FirstOrDefault();
        var relationshipId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;

        if (part.GetPartById(relationshipId) is not ImagePart image)
        {
            inventory.NoteLoss("imagem em formato não suportado");
            return null;
        }

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);

        var node = Node.Of("image")
            .With("src", $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}");

        var extent = drawing.Descendants<Drawing.Wordprocessing.Extent>().FirstOrDefault();
        if (extent?.Cx is not null)
        {
            // EMU → pixels CSS: 914400 EMU por polegada, 96 px por polegada.
            node.With("width", (int)Math.Round(extent.Cx.Value * 96.0 / 914400));
        }

        return node;
    }

    private string? HyperlinkTargetOf(Hyperlink link)
    {
        var id = link.Id?.Value;
        if (string.IsNullOrEmpty(id)) return null;

        try
        {
            return part.HyperlinkRelationships
                .FirstOrDefault(relationship => relationship.Id == id)?.Uri.ToString();
        }
        catch (UriFormatException)
        {
            return null;
        }
    }

    // --- tabelas ------------------------------------------------------------

    private Node ReadTable(Table table)
    {
        var rows = new List<Node>();

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = new List<Node>();

            foreach (var cell in row.Elements<TableCell>())
            {
                var contents = cell.Elements<Paragraph>().Select(ReadParagraph).ToList();
                if (contents.Count == 0) contents.Add(Node.Of("paragraph"));

                var node = Node.Of("tableCell");
                node.Content = contents;

                var span = cell.TableCellProperties?.GridSpan?.Val?.Value;
                if (span is > 1) node.With("colspan", span.Value);

                cells.Add(node);
            }

            var rowNode = Node.Of("tableRow");
            rowNode.Content = cells;
            rows.Add(rowNode);
        }

        var tableNode = Node.Of("table");
        tableNode.Content = rows;
        return tableNode;
    }
}

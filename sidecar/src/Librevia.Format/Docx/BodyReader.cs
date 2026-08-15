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
    private readonly StyleResolver _styles = new(part);
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
        var direct = paragraph.ParagraphProperties;

        // Formatação efetiva: padrões do documento, estilo e formatação direta.
        // Ler só a direta é o que fazia um documento cheio de estilos abrir
        // praticamente sem formatação.
        var (effective, inheritedRun) = _styles.Resolve(direct);

        var content = ReadInline(paragraph, inheritedRun);

        // Uma quebra de página sozinha no parágrafo é o nó `pageBreak`, não um
        // parágrafo vazio com uma quebra dentro.
        if (content.Count == 1 && content[0].Type == "pageBreak")
        {
            return content[0];
        }

        var node = HeadingLevelOf(direct) is { } level
            ? Node.Of("heading").With("level", level)
            : Node.Of("paragraph");

        // O identificador do estilo viaja junto para que a gravação de um
        // parágrafo editado continue apontando o estilo original.
        if (direct?.ParagraphStyleId?.Val?.Value is { Length: > 0 } styleId)
        {
            node.With("styleId", styleId);
        }

        var alignment = AlignmentOf(effective);

        // Tabulações no começo da linha são um posicionador, não texto: o autor
        // alinha à esquerda e usa `Tab` para cair numa parada centralizada. Ver
        // TabAlignmentOf.
        var viaTabs = TabAlignmentOf(effective, content);
        if (viaTabs is not null)
        {
            alignment = viaTabs;
            while (content.Count > 0 && content[0].Type == "text" && content[0].Text == "\t")
            {
                content.RemoveAt(0);
            }
        }

        if (alignment is not null) node.With("textAlign", alignment);

        var indent = IndentOf(effective);
        if (indent > 0) node.With("indent", indent);

        // O fundo do parágrafo é o que transforma `Heading1` numa barra
        // colorida neste corpus — sem ele o título vira texto solto.
        if (ShadingOf(effective) is { } background) node.With("background", background);

        var spacing = effective.SpacingBetweenLines;
        if (TwipsToPt(spacing?.Before?.Value) is { } before) node.With("spaceBefore", before);
        if (TwipsToPt(spacing?.After?.Value) is { } after) node.With("spaceAfter", after);
        if (LineHeightOf(spacing) is { } lineHeight) node.With("lineHeight", lineHeight);

        // A fonte **do bloco**, e não só a dos runs. A altura da linha nasce da
        // fonte do próprio elemento: sem isto, um parágrafo de 10 pt dentro de
        // um bloco que o CSS declara com 12 pt continua ocupando 12 pt de
        // altura — e um `Heading1` de 10 pt vira uma barra alta demais, porque
        // o editor desenha títulos em 22 pt.
        if (FontOf(inheritedRun) is { } font) node.With("fontFamily", font);
        if (FontSizeOf(inheritedRun) is { } size) node.With("fontSize", size);

        // "Manter com o próximo": o parágrafo não fica sozinho no pé da página.
        // É o que faz um rótulo descer junto com a imagem que ele apresenta —
        // e sem ler isto a quebra estimada cai um bloco depois da real.
        if (RunReader.IsOn(effective.KeepNext)) node.With("keepNext", true);

        node.Content = content.Count == 0 ? null : content;
        return node;
    }

    private static string? FontOf(RunProperties properties)
    {
        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        return string.IsNullOrWhiteSpace(font) ? null : font;
    }

    /// <summary>`w:sz` vem em meios-pontos: 20 significa 10 pt.</summary>
    private static string? FontSizeOf(RunProperties properties)
    {
        var value = properties.FontSize?.Val?.Value;
        if (!double.TryParse(value, out var halfPoints) || halfPoints <= 0) return null;
        var points = halfPoints / 2;
        return points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";
    }

    /// <summary>Fundo do parágrafo, quando é cor de verdade.</summary>
    private static string? ShadingOf(ParagraphProperties properties)
    {
        var fill = properties.Shading?.Fill?.Value;
        if (string.IsNullOrWhiteSpace(fill)) return null;
        if (fill.Equals("auto", StringComparison.OrdinalIgnoreCase)) return null;
        // "FFFFFF" explícito é branco de verdade; só "auto" significa "sem cor".
        return "#" + fill.TrimStart('#').ToLowerInvariant();
    }

    /// <summary>
    /// Espaçamento em twips → pontos, que é a unidade da interface.
    /// </summary>
    /// <remarks>
    /// Zero **explícito** é preservado, e não tratado como ausente: "sem espaço
    /// antes" é uma instrução do documento. Descartá-lo deixaria a margem
    /// padrão do editor reaparecer, e o texto sairia mais arejado que no Word.
    /// </remarks>
    private static double? TwipsToPt(string? twips) =>
        int.TryParse(twips, out var value) && value >= 0 ? Math.Round(value / 20.0, 1) : null;

    /// <summary>
    /// Entrelinha. `w:line` com regra `auto` vem em 240-avos: 271 significa
    /// 1,13 vez a altura da linha.
    /// </summary>
    private static double? LineHeightOf(SpacingBetweenLines? spacing)
    {
        if (spacing?.Line?.Value is not { } line || !int.TryParse(line, out var value) || value <= 0) return null;

        var rule = spacing.LineRule?.Value;
        if (rule is not null && rule != LineSpacingRuleValues.Auto) return null;

        var factor = Math.Round(value / 240.0, 2);
        return factor is > 0.5 and < 4 ? factor : null;
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

    /// <summary>
    /// Alinhamento que o autor obteve com tabulações, e não com `w:jc`.
    /// </summary>
    /// <remarks>
    /// No corpus real, o primeiro título de cada documento vem assim: `w:jc` em
    /// `left`, três tabulações e uma **parada de tabulação centralizada** no
    /// meio da coluna. No Word o texto é centralizado naquela parada; no HTML a
    /// tabulação vira espaço em branco que colapsa, e o título encosta à
    /// esquerda enquanto os títulos vizinhos — que usam `w:jc` de verdade —
    /// aparecem centralizados.
    ///
    /// Reproduzir paradas de tabulação em HTML exigiria medir texto e posicionar
    /// à mão. Esta é a aproximação honesta: **só** dispara quando a linha
    /// *começa* com tabulação, então "esquerda [tab] centro" continua intacto.
    /// </remarks>
    private static string? TabAlignmentOf(ParagraphProperties properties, List<Node> content)
    {
        if (content.Count == 0 || content[0].Type != "text" || content[0].Text != "\t") return null;

        var stops = properties.Tabs?.Elements<TabStop>().ToList();
        if (stops is null || stops.Count == 0) return null;

        if (stops.Any(stop => stop.Val is not null && stop.Val.Value == TabStopValues.Center)) return "center";
        if (stops.Any(stop => stop.Val is not null && stop.Val.Value == TabStopValues.Right)) return "right";
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

    private List<Node> ReadInline(
        OpenXmlElement container,
        RunProperties inherited,
        string? hyperlink = null)
    {
        var nodes = new List<Node>();

        foreach (var element in container.ChildElements)
        {
            switch (element)
            {
                case Run run:
                    nodes.AddRange(ReadRun(run, inherited, hyperlink));
                    break;

                case Hyperlink link:
                    nodes.AddRange(ReadInline(link, inherited, HyperlinkTargetOf(link) ?? hyperlink));
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
                    nodes.AddRange(ReadInline(inserted, inherited, hyperlink));
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

    private IEnumerable<Node> ReadRun(Run run, RunProperties inherited, string? hyperlink)
    {
        var marks = RunReader.MarksOf(_styles.ResolveRun(inherited, run.RunProperties), hyperlink);

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

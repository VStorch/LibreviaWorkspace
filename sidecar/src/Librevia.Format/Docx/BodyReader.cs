using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace Librevia.Format.Docx;

/// <summary>Um bloco de primeiro nível do corpo, com identidade.</summary>
/// <param name="Oid">Id estável na ordem do documento: b1, b2, …</param>
/// <param name="Source">O elemento original, para gravar de volta sem tocar.</param>
/// <param name="Extracted">O que o editor vê.</param>
public sealed record Block(string Oid, OpenXmlElement Source, Node Extracted)
{
    /// <summary>
    /// O que o corpo guarda **antes** do bloco e não é bloco — um marcador solto
    /// depois de uma tabela, um controle de conteúdo desconhecido. A gravação o
    /// devolve antes do bloco, byte a byte; sem isto ele cairia fora em silêncio.
    /// </summary>
    public List<OpenXmlElement> Leading { get; } = [];

    /// <summary>O mesmo, depois do último bloco do corpo.</summary>
    public List<OpenXmlElement> Trailing { get; } = [];

    /// <summary>O sumário sem controle de conteúdo, cujo campo fecha num dos parágrafos de baixo.</summary>
    public List<OpenXmlElement> Continuation { get; } = [];

    /// <summary>Campo mostrado só pelo resultado: reescrito, vira texto comum.</summary>
    public bool UnrepresentedField { get; init; }
}

/// <summary>
/// Corpo do documento → nós do editor, **de mão única**: alimenta a tela e o PDF,
/// nunca a gravação, que parte do XML original (<see cref="DocxWriter"/>). Por
/// isso um erro aqui é cosmético, e a leitura pode ser tolerante.
/// </summary>
/// <remarks>
/// Cada interruptor desligado reproduz a leitura de antes do recurso, a de
/// referência para o rascunho daquela época (<see cref="DocumentModelDto.BeforeReferences"/>,
/// <c>BeforeSections</c>, <c>BeforeComments</c>, <c>BeforeRevisions</c>,
/// <c>BeforeNotes</c>, <c>BeforeMath</c>).
/// </remarks>
public sealed class BodyReader(
    MainDocumentPart part,
    Inventory inventory,
    bool flatten = false,
    bool references = true,
    bool sections = true,
    bool comments = true,
    bool revisions = true,
    bool notes = true,
    bool math = true)
{
    /// <summary>O <c>w:footnote</c>/<c>w:endnote</c> e os blocos dele.</summary>
    public sealed record NoteRead(OpenXmlElement Source, List<Block> Blocks);

    /// <summary>Pelo endereço (<c>fn:3</c>, <c>en:1</c>). A nota que o corpo não referencia fica como está.</summary>
    public Dictionary<string, NoteRead> Notes { get; } = new(StringComparer.Ordinal);

    /// <summary>Outra instância: o estado do parágrafo em leitura não pode ser o da nota.</summary>
    private BodyReader? _noteReader;

    /// <summary>A parte dona dos relacionamentos: o documento, ou <c>footnotes.xml</c> na leitura de uma nota.</summary>
    private OpenXmlPart _owner = part;

    /// <summary>Nulo no corpo (<c>b1</c>…); <c>fn:3</c> na nota, cujos blocos saem <c>fn:3/p1</c>…</summary>
    private string? _oidPrefix;

    /// <summary>A de cada <c>w:ins</c>/<c>w:del</c> em volta, de fora para dentro.</summary>
    private readonly List<Mark> _revision = [];

    /// <summary>O nome da movimentação de cada `w:moveFrom`/`w:moveTo` — ver <see cref="MoveNamesOf"/>.</summary>
    private Dictionary<OpenXmlElement, string>? _moveNames;

    /// <summary>
    /// As pontas das respostas não viram nó: o editor leva uma âncora por conversa,
    /// igual na abertura e na leitura de referência, e a impressão digital fica estável.
    /// </summary>
    private readonly HashSet<string> _replies = comments
        ? CommentsReader.RepliesOf(part).Values.SelectMany(ids => ids).ToHashSet(StringComparer.Ordinal)
        : [];

    /// <summary>O comentário sem <c>w:commentRangeEnd</c> é de ponto: a referência vira o <c>commentEnd</c>.</summary>
    private readonly HashSet<string> _rangeEnds = comments
        ? new OpenXmlElement?[] { part.Document, part.FootnotesPart?.Footnotes, part.EndnotesPart?.Endnotes }
            .SelectMany(root => root?.Descendants<CommentRangeEnd>() ?? [])
            .Select(end => end.Id?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal)
        : [];

    /// <summary>Na caixa de texto a âncora não vira nó: a caixa não é reescrita do modelo.</summary>
    private int _textBoxDepth;

    /// <summary>O mesmo id que <see cref="PageReader"/> dá à seção.</summary>
    private Dictionary<SectionProperties, string> _sectionIds = new(ReferenceEqualityComparer.Instance);

    private readonly NumberingReader _numbering = new(part);
    private readonly StyleResolver _styles = new(part);
    private readonly FontTable _fonts = new(part);
    private int _nextId = 1;

    /// <summary>
    /// Decide se uma caixa de texto abre linha nova: as caixas de um parágrafo estão
    /// em <c>w:r</c> diferentes, e nenhuma vê o que a anterior escreveu.
    /// </summary>
    private bool _paragraphHasContent;

    /// <summary>
    /// Só no parágrafo solto do corpo, que o CSS dos estilos desenha. Lista, célula,
    /// caixa e rascunho antigo levam a formatação inteira.
    /// </summary>
    private bool _directRuns;

    /// <summary>Descobertos na leitura de linha, mas do **bloco**: anexados ao nó do parágrafo no fim.</summary>
    private readonly List<FloatDto> _paragraphFloats = [];

    /// <summary>Ver <see cref="TopAnchoredFirst"/>.</summary>
    private readonly HashSet<Node> _topAnchored = new(ReferenceEqualityComparer.Instance);

    /// <summary>Reescrever o bloco perderia o campo que não virou nó — ver <see cref="Block.UnrepresentedField"/>.</summary>
    private bool _unrepresentedField;

    /// <summary>
    /// A árvore e a lista de blocos saem juntas porque os blocos **apontam para nós
    /// da árvore**: um item de lista é um <c>w:p</c> no arquivo e um <c>listItem</c> aninhado.
    /// </summary>
    public (List<Node> Content, List<Block> Blocks) Read(Body body)
    {
        if (sections) _sectionIds = PageReader.SectionIds(body);
        return ReadBlocks(body);
    }

    /// <summary>
    /// Achatado, como a lista e a célula. O run do <c>w:footnoteRef</c>, o número do
    /// começo da nota, não vira nó: quem o refaz é a gravação.
    /// </summary>
    private (List<Node> Content, List<Block> Blocks) ReadNote(OpenXmlElement note, string address, OpenXmlPart owner)
    {
        _owner = owner;
        _oidPrefix = address;
        _nextId = 1;
        return ReadBlocks(note);
    }

    private (List<Node> Content, List<Block> Blocks) ReadBlocks(OpenXmlElement container)
    {
        var content = new List<Node>();
        var blocks = new List<Block>();

        // Parágrafos numerados consecutivos viram uma lista só; uma lista aberta por nível.
        var openLists = new List<(Node List, string Kind, int Level, int NumId)>();

        // Ver Block.Leading.
        var loose = new List<OpenXmlElement>();

        void Add(Block block)
        {
            block.Leading.AddRange(loose);
            loose.Clear();
            blocks.Add(block);
        }

        var elements = container.ChildElements.ToList();
        for (var at = 0; at < elements.Count; at++)
        {
            var element = elements[at];

            // Ver ReadTableOfContents.
            if (references && TableOfContentsAt(elements, at) is { } toc)
            {
                openLists.Clear();
                var block = NewBlock(element, toc.Node);
                block.Continuation.AddRange(toc.Continuation);
                Add(block);
                content.Add(toc.Node);
                at += toc.Continuation.Count;
                continue;
            }

            switch (element)
            {
                case Paragraph paragraph:
                {
                    // Só o parágrafo solto no corpo deixa de ser achatado (`.page__content > p`).
                    var numbered = _numbering.ListKindOf(paragraph.ParagraphProperties);
                    var node = ReadParagraph(paragraph, flat: flatten || _oidPrefix is not null || numbered is not null);
                    var list = node.Type == "pageBreak" ? null : numbered;

                    if (list is null)
                    {
                        openLists.Clear();
                        Add(NewBlock(element, node));
                        content.Add(node);
                        break;
                    }

                    var (kind, level) = (list.Kind, list.Level);
                    while (openLists.Count > 0 && openLists[^1].Level > level)
                    {
                        openLists.RemoveAt(openLists.Count - 1);
                    }

                    // No mesmo nível, outra numeração é outra lista, como no Word.
                    if (openLists.Count > 0 && openLists[^1].Level == level &&
                        (openLists[^1].Kind != kind || openLists[^1].NumId != list.NumberingId))
                    {
                        openLists.RemoveAt(openLists.Count - 1);
                    }

                    if (openLists.Count == 0 || openLists[^1].Level < level)
                    {
                        var listNode = Node.Of(kind);
                        listNode.Content = [];
                        var parent = openLists.Count > 0 ? openLists[^1] : default;
                        var depth = openLists.Count;

                        // A marca e o recuo são do nível. O `numId` viaja no nó: sem
                        // ele o escritor gravaria `w:numId w:val="0"`, "sem numeração".
                        listNode.With("numId", list.NumberingId);

                        if (list.Marker is { } marker) listNode.With("marker", marker);

                        // O `start` que o editor materializa em toda lista numerada:
                        // entra na impressão digital do item de fora.
                        if (kind == "orderedList") listNode.With("start", 1);
                        if (list.IndentMm is { } indent) listNode.With("indentMm", indent);
                        if (list.HangingMm is { } hanging) listNode.With("hangingMm", hanging);

                        // A definição só onde a numeração muda.
                        if (depth == 0 || parent.NumId != list.NumberingId)
                        {
                            listNode.With("numbering", list.Definition);
                        }

                        // A lista que começa no nível 2 mora no topo da árvore: sem
                        // isto voltaria ao arquivo no nível 0.
                        if (level != (depth == 0 ? 0 : parent.Level + 1)) listNode.With("level", level);

                        if (depth > 0)
                        {
                            var parentItems = parent.List.Content!;
                            (parentItems[^1].Content ??= []).Add(listNode);
                        }
                        else
                        {
                            content.Add(listNode);
                        }

                        openLists.Add((listNode, kind, level, list.NumberingId));
                    }

                    // No `listItem`: é o item que corresponde a um `w:p`.
                    var item = Node.Of("listItem");
                    item.Content = [node];
                    openLists[^1].List.Content!.Add(item);
                    CarrySpacing(openLists[^1].List, node);
                    Add(NewBlock(element, item));

                    // A seção termina no item, como no Word: o seguinte abre outra lista.
                    if (node.Attrs?.ContainsKey("sectionBreak") == true) openLists.Clear();
                    break;
                }

                case Table table:
                {
                    openLists.Clear();
                    var node = ReadTable(table);
                    Add(NewBlock(element, node));
                    content.Add(node);
                    break;
                }

                case SectionProperties:
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    loose.Add(element);
                    break;
            }
        }

        // O que sobrou depois do último bloco fica com ele.
        if (blocks.Count > 0) blocks[^1].Trailing.AddRange(loose);

        if (content.Count == 0) content.Add(Node.Of("paragraph"));

        return (content, blocks);
    }

    /// <summary>
    /// Na árvore a lista é um elemento e receberia o espaçamento do editor: num
    /// documento com seis listas, quinze milímetros a mais. Antes do primeiro item
    /// vale o espaço de antes dele; depois do último, o de depois, como no Word.
    /// </summary>
    private static void CarrySpacing(Node list, Node paragraph)
    {
        if (paragraph.Attrs is not { } attrs) return;

        if (list.Attrs?.ContainsKey("spaceBefore") != true && attrs.TryGetValue("spaceBefore", out var before))
        {
            list.With("spaceBefore", before?.DeepClone());
        }

        if (attrs.TryGetValue("spaceAfter", out var after)) list.With("spaceAfter", after?.DeepClone());
    }

    private Block NewBlock(OpenXmlElement source, Node extracted)
    {
        var oid = _oidPrefix is null ? "b" + _nextId++ : $"{_oidPrefix}/p{_nextId++}";
        extracted.With("oid", oid);
        var block = new Block(oid, source, extracted) { UnrepresentedField = _unrepresentedField };
        _unrepresentedField = false;
        return block;
    }


    private Node ReadParagraph(Paragraph paragraph, bool flat)
    {
        var direct = paragraph.ParagraphProperties;

        // Formatação efetiva: padrões, estilo e direta.
        var (effective, inheritedRun) = _styles.Resolve(direct);

        _paragraphHasContent = false;
        _paragraphFloats.Clear();
        _topAnchored.Clear();
        var outer = _directRuns;
        _directRuns = !flat;
        var content = ReadInline(paragraph, inheritedRun);
        _directRuns = outer;

        // Uma quebra de página sozinha é o nó `pageBreak`, **menos** quando o
        // parágrafo ancora algo: o objeto cairia na folha de baixo. Aí a quebra
        // vira propriedade do parágrafo.
        if (content.Count == 1 && content[0].Type == "pageBreak" && _paragraphFloats.Count == 0)
        {
            return content[0];
        }

        // A quebra no meio do parágrafo (`w:br w:type="page"` num `w:r`) vira
        // propriedade do bloco: como nó em posição de linha seria inválido, e no
        // HTML desalinharia os índices entre tela e papel. Exato quando a quebra
        // encerra o parágrafo; aproximado quando há texto depois dela.
        var breakAfter = content.RemoveAll(child => child.Type == "pageBreak") > 0;
        var columnBreakAfter = content.RemoveAll(child => child.Type == "columnBreak") > 0;
        TopAnchoredFirst(content);

        var node = (HeadingLevelOf(direct) ?? _styles.HeadingLevelByName(direct?.ParagraphStyleId?.Val?.Value)) is { } level
            ? Node.Of("heading").With("level", level)
            : Node.Of("paragraph");

        if (breakAfter) node.With("breakAfter", true);
        if (columnBreakAfter) node.With("columnBreakAfter", true);

        WithFloats(node);

        // Para o parágrafo editado continuar apontando o estilo original.
        if (direct?.ParagraphStyleId?.Val?.Value is { Length: > 0 } styleId)
        {
            node.With("styleId", styleId);
        }

        var alignment = AlignmentOf(effective);

        // Tabulações no começo da linha posicionam, não são texto — ver TabAlignmentOf.
        var viaTabs = TabAlignmentOf(effective, content);
        if (viaTabs is not null)
        {
            alignment = viaTabs;
            while (content.Count > 0 && content[0].Type == "text" && content[0].Text == "\t")
            {
                content.RemoveAt(0);
            }
        }

        if (flat) Flattened(node, alignment, effective, direct, inheritedRun);
        else DirectOnly(node, viaTabs, effective, direct, inheritedRun);

        // O parágrafo vazio que guarda o `w:sectPr` **é** a marca de seção, sem
        // altura, como no LibreOffice. Continua no modelo: a gravação o devolve.
        if (direct?.SectionProperties is not null && content.Count == 0)
        {
            node.With("sectionMark", true);
        }

        // A configuração mora fora dos nós, em `sections`; o id é o elo.
        if (direct?.SectionProperties is { } marked && _sectionIds.TryGetValue(marked, out var sectionId))
        {
            node.With("sectionBreak", sectionId);
        }

        // A marca de parágrafo revisada: aceitar a exclusão junta este ao seguinte.
        if (revisions && Revisions.BlockRevisionOf(direct?.ParagraphMarkRunProperties) is { } markRevision)
        {
            node.With("markRevision", markRevision);
        }

        node.Content = content.Count == 0 ? null : content;
        return node;
    }
    /// <summary>
    /// Padrões, estilo e direta, achatados: onde as regras dos estilos não chegam
    /// (lista, célula) e no modo <c>flatten</c>, a leitura de referência do rascunho
    /// achatado.
    /// </summary>
    private void Flattened(
        Node node,
        string? alignment,
        ParagraphProperties effective,
        ParagraphProperties? direct,
        RunProperties inheritedRun)
    {
        if (alignment is not null) node.With("textAlign", alignment);


        // Zero, sempre: o recuo do arquivo vem em milímetros logo abaixo, e o
        // editor devolve `indent` em todo parágrafo, então omiti-lo faria o bloco
        // parecer mudado.
        node.With("indent", 0);

        // Em milímetros, como o arquivo declara.
        Measure(node, "indentMm", effective.Indentation?.Left?.Value);
        Measure(node, "indentRightMm", effective.Indentation?.Right?.Value);

        // `w:firstLine` empurra e `w:hanging` puxa: o mesmo `text-indent` do CSS.
        var firstLine = TwipsToMm(effective.Indentation?.FirstLine?.Value);
        var hanging = TwipsToMm(effective.Indentation?.Hanging?.Value);
        if (firstLine is > 0) node.With("firstLineMm", firstLine.Value);
        else if (hanging is > 0) node.With("firstLineMm", -hanging.Value);

        // É o fundo que faz o `Heading1` do corpus virar barra colorida.
        if (ShadingOf(effective) is { } background) node.With("background", background);

        // **Sempre** escritos: silêncio no arquivo é zero, e não o padrão do editor.
        var spacing = effective.SpacingBetweenLines;
        node.With("spaceBefore", TwipsToPt(spacing?.Before?.Value) ?? 0);
        node.With("spaceAfter", TwipsToPt(spacing?.After?.Value) ?? 0);
        // A fonte que mede a linha é a da marca do parágrafo.
        var markFont = _styles.ResolveMark(inheritedRun, direct).RunFonts?.Ascii?.Value;
        node.With("lineHeight", LineHeightOf(spacing, LineMetrics.Of(markFont)));

        // A fonte **do bloco**, da marca de parágrafo (`w:pPr/w:rPr`): a altura da
        // linha nasce da fonte do elemento, e é a marca que dá altura ao
        // parágrafo vazio, como no Word.
        var mark = _styles.ResolveMark(inheritedRun, direct);
        if (FontOf(mark) is { } font) node.With("fontFamily", font);
        if (FontSizeOf(mark) is { } size) node.With("fontSize", size);

        if (RunReader.IsOn(effective.KeepNext)) node.With("keepNext", true);

        if (RunReader.IsOn(effective.KeepLines)) node.With("keepLines", true);

        // Ligado quando o arquivo cala, como no Word: só o desligado se diz.
        if (effective.WidowControl is { } widow && !RunReader.IsOn(widow)) node.With("widowControl", false);
    }

    /// <summary>
    /// Só a formatação **direta**; o herdado vem do CSS dos estilos.
    /// </summary>
    /// <remarks>
    /// O arquivo declara a propriedade no <c>w:pPr</c> (ou na marca, para a fonte)?
    /// O bloco leva o valor **efetivo** dela, como no achatamento; se cala, o bloco
    /// cala. Zero declarado é declaração. Duas exceções: o alinhamento por
    /// tabulação (<see cref="TabAlignmentOf"/>) e a entrelinha com a fonte da marca
    /// direta, que se mede sobre a fonte dela.
    /// </remarks>
    private void DirectOnly(
        Node node,
        string? viaTabs,
        ParagraphProperties effective,
        ParagraphProperties? direct,
        RunProperties inheritedRun)
    {
        if (viaTabs is not null) node.With("textAlign", viaTabs);
        else if (direct?.Justification is not null && AlignmentOf(effective) is { } alignment)
        {
            node.With("textAlign", alignment);
        }

        // Zero pelo mesmo motivo do achatamento.
        node.With("indent", 0);

        // Recuo negativo sai como zero: o bloco não desenha para fora da margem.
        var indentation = direct?.Indentation;
        if (indentation?.Left is not null) Declared(node, "indentMm", effective.Indentation?.Left?.Value);
        if (indentation?.Right is not null) Declared(node, "indentRightMm", effective.Indentation?.Right?.Value);
        if (indentation?.FirstLine is not null || indentation?.Hanging is not null)
        {
            var firstLine = TwipsToMm(effective.Indentation?.FirstLine?.Value);
            var hanging = TwipsToMm(effective.Indentation?.Hanging?.Value);
            node.With("firstLineMm", firstLine is > 0 ? firstLine.Value : hanging is > 0 ? -hanging.Value : 0);
        }

        if (direct?.Shading is not null)
        {
            // `w:shd` sem cor sobre estilo com fundo apaga o fundo: senão a regra o pintaria.
            if (ShadingOf(effective) is { } background) node.With("background", background);
            else if (ShadingOf(_styles.StyleParagraphOf(direct)) is not null) node.With("background", "transparent");
        }

        var spacing = effective.SpacingBetweenLines;
        var declared = direct?.SpacingBetweenLines;
        if (declared?.Before is not null) node.With("spaceBefore", TwipsToPt(spacing?.Before?.Value) ?? 0);
        if (declared?.After is not null) node.With("spaceAfter", TwipsToPt(spacing?.After?.Value) ?? 0);

        var mark = _styles.ResolveMark(inheritedRun, direct);
        var lineHeight = LineHeightOf(spacing, LineMetrics.Of(mark.RunFonts?.Ascii?.Value), 4);
        if (declared?.Line is not null ||
            lineHeight != LineHeightOf(spacing, LineMetrics.Of(inheritedRun.RunFonts?.Ascii?.Value), 4))
        {
            node.With("lineHeight", lineHeight);
        }

        var markDirect = direct?.ParagraphMarkRunProperties;
        if (markDirect?.GetFirstChild<RunFonts>() is { } fonts &&
            (fonts.Ascii is not null || fonts.HighAnsi is not null) &&
            FontOf(mark) is { } font)
        {
            node.With("fontFamily", font);
        }

        if (markDirect?.GetFirstChild<FontSize>() is not null && FontSizeOf(mark) is { } size)
        {
            node.With("fontSize", size);
        }

        // `w:keepNext w:val="0"` existe para desfazer o do estilo.
        if (direct?.KeepNext is not null) node.With("keepNext", RunReader.IsOn(effective.KeepNext));
        if (direct?.KeepLines is not null) node.With("keepLines", RunReader.IsOn(effective.KeepLines));
        if (direct?.WidowControl is not null) node.With("widowControl", RunReader.IsOn(effective.WidowControl));
    }

    /// <summary>Uma medida declarada: zero conta, negativo vira zero.</summary>
    private static void Declared(Node node, string name, string? twips)
    {
        if (TwipsToMm(twips) is { } value) node.With(name, Math.Max(0, value));
    }

    private string? FontOf(RunProperties properties)
    {
        var font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
        return string.IsNullOrWhiteSpace(font) ? null : _fonts.Stack(font);
    }

    /// <summary><c>w:sz</c> vem em meios-pontos.</summary>
    private static string? FontSizeOf(RunProperties properties)
    {
        var value = properties.FontSize?.Val?.Value;
        if (!double.TryParse(value, out var halfPoints) || halfPoints <= 0) return null;
        var points = halfPoints / 2;
        return points == Math.Floor(points)
            ? $"{(int)points}pt"
            : points.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";
    }

    private static string? ShadingOf(ParagraphProperties properties)
    {
        var fill = properties.Shading?.Fill?.Value;
        if (string.IsNullOrWhiteSpace(fill)) return null;
        if (fill.Equals("auto", StringComparison.OrdinalIgnoreCase)) return null;
        // "FFFFFF" é branco de verdade; só "auto" é "sem cor".
        return "#" + fill.TrimStart('#').ToLowerInvariant();
    }

    /// <summary>Twips → pontos. Zero **explícito** é "sem espaço antes", e não ausência.</summary>
    private static double? TwipsToPt(string? twips) =>
        int.TryParse(twips, out var value) && value >= 0 ? Math.Round(value / 20.0, 1) : null;

    /// <summary>
    /// Entrelinha em CSS. <c>w:line</c> com regra <c>auto</c> vem em 240-avos de **vez a
    /// altura natural**, e não do tamanho da fonte. Sai número sempre que se sabe a
    /// fonte, inclusive no simples, porque o Chromium arredonda <c>normal</c> para
    /// pixel inteiro; com fonte desconhecida, <c>normal</c>. <c>exact</c> e
    /// <c>atLeast</c> viram pontos.
    /// </summary>
    /// <param name="decimals">
    /// Duas no bloco achatado, como sempre foi; quatro no que leva só o direto, a
    /// grade de 240-avos em que <see cref="StyleReader"/> também entrega o do estilo.
    /// </param>
    private static string LineHeightOf(SpacingBetweenLines? spacing, double? natural, int decimals = 2)
    {
        var rule = spacing?.LineRule?.Value;
        var declared = spacing?.Line?.Value;

        if (declared is not null && int.TryParse(declared, out var value) && value > 0)
        {
            if (rule is not null && rule != LineSpacingRuleValues.Auto)
            {
                return Math.Round(value / 20.0, 1)
                    .ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "pt";
            }

            var factor = Math.Round(value / 240.0, decimals);

            // Fora dessa faixa é lixo do arquivo.
            if (factor is > 0.5 and < 4) return Multiple(factor, natural);
        }

        return Multiple(1, natural);
    }

    /// <summary>
    /// Com fonte desconhecida e múltiplo declarado, o palpite de 1,15: a altura de
    /// quase toda fonte latina e das substitutas do LibreOffice.
    /// </summary>
    private static string Multiple(double factor, double? natural)
    {
        if (natural is not null) return Text(Math.Round(factor * natural.Value, 4));
        return factor == 1 ? "normal" : Text(Math.Round(factor * 1.1499, 4));
    }

    private static string Text(double value) =>
        value.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>
    /// <c>Heading1</c>, <c>heading 1</c> e o <c>Ttulo1</c> do LibreOffice, sem acento
    /// porque o id do estilo não os aceita.
    /// </summary>
    private static int? HeadingLevelOf(ParagraphProperties? properties) =>
        HeadingLevelOfStyle(properties?.ParagraphStyleId?.Val?.Value);

    /// <summary>Para <see cref="ParagraphFormat"/> não trocar o <c>Ttulo1</c> do documento por um <c>Heading1</c>.</summary>
    internal static int? HeadingLevelOfStyle(string? style)
    {
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
    /// No corpus, o primeiro título vem com <c>w:jc</c> à esquerda, tabulações e uma
    /// parada centralizada: no HTML a tabulação colapsaria. Só dispara quando a
    /// linha **começa** com tabulação.
    /// </summary>
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

    private Node WithFloats(Node node)
    {
        if (_paragraphFloats.Count > 0)
        {
            node.With("floats", JsonSerializer.SerializeToNode(_paragraphFloats, DocxJson.Options));
        }

        return node;
    }

    private static void Measure(Node node, string name, string? twips)
    {
        if (TwipsToMm(twips) is { } value and > 0) node.With(name, value);
    }

    private static double? TwipsToMm(string? twips) =>
        int.TryParse(twips, out var value) ? Math.Round(value * 25.4 / 1440, 2) : null;


    private List<Node> ReadInline(
        OpenXmlElement container,
        RunProperties inherited,
        string? hyperlink = null)
    {
        var nodes = new List<Node>();
        var children = container.ChildElements.ToList();

        for (var index = 0; index < children.Count; index++)
        {
            var element = children[index];

            // O campo inteiro vira nó (ReadField); o que não couber segue run a run e vai ao inventário.
            if (references && element is Run begin && IsFieldBegin(begin) &&
                ReadField(children, index, inherited, hyperlink) is { } field)
            {
                nodes.Add(field.Node);
                index = field.Last;
                continue;
            }

            // A marca própria (`w:customMarkFollows`) no run seguinte é a marca, e não texto.
            if (notes && _textBoxDepth == 0 && element is Run noted && NoteReferenceOf(noted) is { } noteReference)
            {
                var read = ReadRun(noted, inherited, hyperlink).ToList();
                if (read.LastOrDefault(node => node.Type == "noteRef") is { } noteRef &&
                    IsCustomMark(noteReference) && noteRef.Attrs?.ContainsKey("mark") != true &&
                    index + 1 < children.Count && children[index + 1] is Run follower &&
                    string.Concat(follower.Elements<Text>().Select(text => text.Text)) is { Length: > 0 } mark)
                {
                    noteRef.With("mark", mark);
                    index++;
                }

                nodes.AddRange(read);
                continue;
            }

            if (references && element is SimpleField simple &&
                ReadSimpleField(simple, inherited, hyperlink) is { } plain)
            {
                nodes.Add(plain);
                continue;
            }

            switch (element)
            {
                // As pontas viram nós sem largura, como as do marcador; a resposta não (`_replies`).
                case CommentRangeStart start when comments && _textBoxDepth == 0:
                    if (start.Id?.Value is { } startId && !_replies.Contains(startId))
                    {
                        nodes.Add(Node.Of("commentStart").With("cid", startId));
                    }

                    break;

                case CommentRangeEnd end when comments && _textBoxDepth == 0:
                    if (end.Id?.Value is { } endId && !_replies.Contains(endId))
                    {
                        nodes.Add(Node.Of("commentEnd").With("cid", endId));
                    }

                    break;

                // A gravação refaz o run da referência com o `commentEnd`; só o comentário de ponto ganha o fim aqui.
                case Run reference when comments && ReferenceOnly(reference) is { } referenced:
                    if (_textBoxDepth == 0 && !_replies.Contains(referenced) && !_rangeEnds.Contains(referenced))
                    {
                        nodes.Add(Node.Of("commentEnd").With("cid", referenced));
                    }

                    break;

                case Run run:
                    nodes.AddRange(ReadRun(run, inherited, hyperlink));
                    break;

                case Hyperlink link:
                    nodes.AddRange(ReadInline(link, inherited, HyperlinkTargetOf(link) ?? hyperlink));
                    break;

                // As pontas viram nós sem largura e sobrevivem à edição do
                // parágrafo; os ocultos entram também, porque o Word os cita.
                case BookmarkStart start when references:
                    nodes.Add(Node.Of("bookmarkStart")
                        .With("name", start.Name?.Value ?? string.Empty)
                        .With("bid", start.Id?.Value ?? string.Empty));
                    break;

                case BookmarkEnd end when references:
                    nodes.Add(Node.Of("bookmarkEnd").With("bid", end.Id?.Value ?? string.Empty));
                    break;

                // O campo simples que o nó não representa (ReadSimpleField).
                case SimpleField:
                    inventory.NoteInvisible(Inventory.Fields);
                    _unrepresentedField = true;
                    break;

                case ParagraphProperties:
                case BookmarkStart:
                case BookmarkEnd:
                case ProofError:
                    break;

                // O da caixa de texto e o do rascunho anterior aos comentários
                // voltam pelo XML original: invisibilidade, não perda.
                case CommentRangeStart:
                case CommentRangeEnd:
                    inventory.NoteInvisible(Inventory.Comments);
                    break;

                // O trecho leva a marca da revisão, em qualquer ordem com o link.
                case InsertedRun or DeletedRun or MoveFromRun or MoveToRun when revisions:
                    _revision.Add(Revisions.MarkOf(element, MoveNameOf(element)));
                    nodes.AddRange(ReadInline(element, inherited, hyperlink));
                    _revision.RemoveAt(_revision.Count - 1);
                    break;

                // O rascunho anterior às revisões (`BeforeRevisions`) lê o
                // inserido como texto comum.
                case InsertedRun inserted:
                    nodes.AddRange(ReadInline(inserted, inherited, hyperlink));
                    break;

                case DeletedRun:
                    break;

                // O OMML volta ao arquivo; o MathML é o que a tela desenha.
                case OfficeMath or MathParagraph when math && _textBoxDepth == 0:
                    nodes.Add(ReadMath(element));
                    break;

                // A de dentro de uma caixa volta com o XML da caixa.
                case OfficeMath or MathParagraph when math:
                    inventory.NoteInvisible(Inventory.Equations);
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }

        return nodes;
    }

    /// <summary>
    /// O <c>omml</c> é a identidade: a impressão digital só vê ele, porque o MathML
    /// muda quando a conversão melhora. O que a conversão não desenha é
    /// invisibilidade, e não perda.
    /// </summary>
    private Node ReadMath(OpenXmlElement element)
    {
        var omml = element.OuterXml;
        var converted = OmmlMath.Convert(omml);
        if (converted.Lossy.Count > 0) inventory.NoteInvisible(Inventory.Equations);

        var node = new Node { Type = "math", Marks = RevisionMarks() }
            .With("omml", omml)
            .With("mathml", converted.MathMl)
            .With("latex", string.Empty)
            .With("display", converted.Display)
            .With("lossy", new JsonArray([.. converted.Lossy.Select(name => (JsonNode?)JsonValue.Create(name))]))
            .With("editable", converted.Lossy.Count == 0);
        if (converted.Jc is { } jc) node.With("jc", jc);
        return node;
    }

    /// <summary>Nulo quando o run traz mais que a referência.</summary>
    internal static string? ReferenceOnly(Run run)
    {
        string? id = null;
        foreach (var child in run.ChildElements)
        {
            if (child is RunProperties) continue;
            if (child is not CommentReference reference || id is not null) return null;
            id = reference.Id?.Value ?? string.Empty;
        }

        return id;
    }

    internal static OpenXmlElement? NoteReferenceOf(Run run) =>
        run.ChildElements.FirstOrDefault(child => child is FootnoteReference or EndnoteReference);

    /// <summary><c>w:customMarkFollows</c>: a marca é o texto que vem depois.</summary>
    internal static bool IsCustomMark(OpenXmlElement reference) =>
        (reference as FootnoteEndnoteReferenceType)?.CustomMarkFollows?.Value == true;

    private Node ReadNoteRef(OpenXmlElement reference)
    {
        var endnote = reference is EndnoteReference;
        var id = (reference as FootnoteEndnoteReferenceType)?.Id?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var node = Node.Of("noteRef")
            .With("kind", endnote ? NotesWriter.Endnote : NotesWriter.Footnote)
            .With("nid", id);
        node.Marks = RevisionMarks();

        var source = id is null ? null : NotesWriter.NoteOf(part, endnote, id);
        if (source is null)
        {
            node.Content = [Node.Of("paragraph")];
            return node;
        }

        var owner = endnote ? (OpenXmlPart)part.EndnotesPart! : part.FootnotesPart!;
        var address = NotesWriter.Address(endnote, id!);
        _noteReader ??= new BodyReader(part, inventory, flatten, references, sections: false, comments, revisions, notes: false, math);
        var (content, blocks) = _noteReader.ReadNote(source, address, owner);
        Notes[address] = new NoteRead(source, blocks);
        node.Content = content;
        return node;
    }

    /// <summary>Mora no <c>w:moveFromRangeStart</c>/<c>w:moveToRangeStart</c>, e não no run.</summary>
    private string? MoveNameOf(OpenXmlElement revision)
    {
        if (revision is not (MoveFromRun or MoveToRun)) return null;
        _moveNames ??= MoveNamesOf(part);
        return _moveNames.GetValueOrDefault(revision);
    }

    private static Dictionary<OpenXmlElement, string> MoveNamesOf(MainDocumentPart part)
    {
        var names = new Dictionary<OpenXmlElement, string>(ReferenceEqualityComparer.Instance);
        var open = new List<(string Id, string Name, bool From)>();
        foreach (var element in part.Document?.Descendants() ?? [])
        {
            switch (element)
            {
                case MoveFromRangeStart start:
                    open.Add((start.Id?.Value ?? string.Empty, start.Name?.Value ?? string.Empty, true));
                    break;
                case MoveToRangeStart start:
                    open.Add((start.Id?.Value ?? string.Empty, start.Name?.Value ?? string.Empty, false));
                    break;
                case MoveFromRangeEnd end:
                    open.RemoveAll(range => range.From && range.Id == end.Id?.Value);
                    break;
                case MoveToRangeEnd end:
                    open.RemoveAll(range => !range.From && range.Id == end.Id?.Value);
                    break;
                case MoveFromRun or MoveToRun:
                    var from = element is MoveFromRun;
                    var range = open.LastOrDefault(candidate => candidate.From == from);
                    if (range.Name is { Length: > 0 } name) names[element] = name;
                    break;
            }
        }

        return names;
    }

    /// <summary>Para o nó que não é texto, como a quebra de linha.</summary>
    private List<Mark>? RevisionMarks() => _revision.Count == 0 ? null : [.. _revision];

    private List<Mark>? MarksOfRun(Run run, RunProperties inherited, string? hyperlink)
    {
        // O herdado vai junto para o "desligado" direto virar marca; o rascunho antigo não as conhece.
        var marks = RunReader.MarksOf(
            _styles.ResolveRun(inherited, run.RunProperties),
            hyperlink,
            _fonts,
            flatten ? null : inherited,
            directOnly: _directRuns);

        // `w:rStyle`: quem desenha é o CSS dos estilos, e a marca o faz voltar ao arquivo.
        if (!flatten && run.RunProperties?.RunStyle?.Val?.Value is { Length: > 0 } characterStyle)
        {
            (marks ??= []).Add(Mark.Of("charStyle", "styleId", characterStyle));
        }

        if (_revision.Count > 0) (marks ??= []).AddRange(_revision);

        return marks;
    }


    private sealed record TableOfContentsRead(Node Node, List<OpenXmlElement> Continuation);

    /// <summary>
    /// Um <c>w:sdt</c> "Table of Contents", como o Word e o LibreOffice gravam, ou os
    /// próprios parágrafos, quando o campo <c>TOC</c> abre no primeiro e fecha num dos
    /// de baixo (<see cref="Block.Continuation"/>).
    /// </summary>
    private TableOfContentsRead? TableOfContentsAt(List<OpenXmlElement> elements, int at)
    {
        if (elements[at] is SdtBlock sdt)
        {
            var inside = sdt.SdtContentBlock?.ChildElements.ToList() ?? [];
            if (inside.Count == 0 || inside.Any(child => child is not Paragraph)) return null;
            var node = ReadTableOfContents([.. inside.Cast<Paragraph>()], sdt: true);
            return node is null ? null : new TableOfContentsRead(node, []);
        }

        if (elements[at] is not Paragraph first || TocBegin(first) is null) return null;

        // Até o parágrafo em que o campo fecha; o sumário não atravessa tabela.
        var depth = 0;
        for (var index = at; index < elements.Count; index++)
        {
            if (elements[index] is not Paragraph paragraph) return null;
            foreach (var mark in paragraph.Descendants<FieldChar>())
            {
                if (mark.FieldCharType?.Value == FieldCharValues.Begin) depth++;
                else if (mark.FieldCharType?.Value == FieldCharValues.End) depth--;
            }

            if (depth > 0) continue;

            var group = elements.Skip(at).Take(index - at + 1).Cast<Paragraph>().ToList();
            var node = ReadTableOfContents(group, sdt: false);
            return node is null ? null : new TableOfContentsRead(node, [.. group.Skip(1)]);
        }

        return null;
    }

    /// <summary>O índice do run que abre o <c>TOC</c> de primeiro nível, o do separador e a instrução.</summary>
    private static (int Begin, int Separate, string Instruction)? TocBegin(Paragraph paragraph)
    {
        var children = paragraph.ChildElements.ToList();
        var begin = children.FindIndex(child => child is Run run && IsFieldBegin(run));
        if (begin < 0) return null;

        var instruction = new System.Text.StringBuilder();
        for (var index = begin + 1; index < children.Count; index++)
        {
            if (children[index] is not Run run) return null;
            if (FieldCharOf(run) is { } mark)
            {
                if (mark.FieldCharType?.Value != FieldCharValues.Separate) return null;
                var text = instruction.ToString();
                return text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
                    ? (begin, index, text)
                    : null;
            }

            foreach (var code in run.Elements<FieldCode>()) instruction.Append(code.Text);
        }

        return null;
    }

    /// <summary>
    /// O campo <c>TOC</c> sai dos parágrafos e vai para o nó; o que sobra são parágrafos
    /// comuns, editáveis à mão como no Word. <c>head</c> conta os parágrafos antes do
    /// campo. O campo que não fecha no último parágrafo não é representado.
    /// </summary>
    private Node? ReadTableOfContents(List<Paragraph> paragraphs, bool sdt)
    {
        var head = paragraphs.FindIndex(paragraph => paragraph.Descendants<FieldChar>().Any());
        if (head < 0 || TocBegin(paragraphs[head]) is not { } begin) return null;

        // Contando os PAGEREF de dentro.
        var depth = 0;
        FieldChar? end = null;
        var endParagraph = -1;
        for (var index = head; index < paragraphs.Count && end is null; index++)
        {
            foreach (var mark in paragraphs[index].Descendants<FieldChar>())
            {
                if (mark.FieldCharType?.Value == FieldCharValues.Begin) depth++;
                else if (mark.FieldCharType?.Value == FieldCharValues.End && --depth == 0)
                {
                    end = mark;
                    endParagraph = index;
                    break;
                }
            }
        }

        if (end is null || endParagraph != paragraphs.Count - 1) return null;
        if (end.Parent is not Run endRun || endRun.Parent != paragraphs[endParagraph] || FieldCharOf(endRun) is null)
        {
            return null;
        }

        var endIndex = paragraphs[endParagraph].ChildElements.ToList().IndexOf(endRun);
        var clones = paragraphs.Select(paragraph => (Paragraph)paragraph.CloneNode(true)).ToList();

        // O fim antes do começo, para o índice valer no sumário de um parágrafo só.
        clones[endParagraph].ChildElements[endIndex].Remove();
        var opening = clones[head].ChildElements.Skip(begin.Begin).Take(begin.Separate - begin.Begin + 1).ToList();
        foreach (var piece in opening) piece.Remove();

        var content = new List<Node>();
        foreach (var clone in clones)
        {
            var node = ReadParagraph(clone, flat: flatten);
            if (node.Type is not ("paragraph" or "heading")) return null;
            content.Add(node);
        }

        var toc = Node.Of("tableOfContents")
            .With("instr", begin.Instruction)
            .With("head", head)
            .With("sdt", sdt);
        toc.Content = content;
        return toc;
    }


    /// <summary>Só o <c>w:fldChar</c> de início.</summary>
    private static bool IsFieldBegin(Run run) =>
        FieldCharOf(run) is { } mark && mark.FieldCharType?.Value == FieldCharValues.Begin;

    /// <summary>Quando é a única coisa que o run carrega.</summary>
    private static FieldChar? FieldCharOf(Run run)
    {
        var content = run.ChildElements.Where(child => child is not RunProperties).ToList();
        return content is [FieldChar only] ? only : null;
    }

    private sealed record ReadFieldResult(Node Node, int Last);

    /// <summary>
    /// Um campo complexo inteiro como nó <c>field</c>, para sobreviver à edição do
    /// parágrafo. Só o caso sem perda: um contêiner, um run por peça, resultado de
    /// texto e tabulação, sem campo aninhado. Fora disso devolve nulo, e o campo é
    /// lido pelo resultado. A formatação é a do primeiro run do resultado.
    /// </summary>
    private ReadFieldResult? ReadField(List<OpenXmlElement> siblings, int start, RunProperties inherited, string? hyperlink)
    {
        // O que o nó não carrega (`w:ffData`, `w:fldData`, `w:fldLock`, `w:dirty`)
        // fica fora dele: regravado pelo nó, o campo voltaria sem isso.
        if (FieldCharOf((Run)siblings[start]) is { } opening &&
            (opening.HasChildren || opening.FieldLock is not null || opening.Dirty is not null))
        {
            return null;
        }

        var instruction = new System.Text.StringBuilder();
        var result = new System.Text.StringBuilder();
        Run? formatted = null;
        var separated = false;

        for (var index = start + 1; index < siblings.Count; index++)
        {
            if (siblings[index] is ProofError) continue;
            if (siblings[index] is not Run run) return null;

            foreach (var child in run.ChildElements)
            {
                switch (child)
                {
                    case RunProperties:
                    case LastRenderedPageBreak:
                        break;

                    case FieldChar mark when mark.FieldCharType?.Value == FieldCharValues.Separate && !separated:
                        separated = true;
                        break;

                    case FieldChar mark when mark.FieldCharType?.Value == FieldCharValues.End:
                        if (FieldCharOf(run) is null) return null;
                        var node = Node.Of("field")
                            .With("instr", instruction.ToString())
                            .With("result", result.ToString());
                        node.Marks = MarksOfRun(formatted ?? (Run)siblings[start], inherited, hyperlink);
                        if (result.Length > 0) _paragraphHasContent = true;
                        return new ReadFieldResult(node, index);

                    case FieldCode code when !separated:
                        instruction.Append(code.Text);
                        break;

                    case DeletedFieldCode code when !separated && revisions:
                        instruction.Append(code.Text);
                        break;

                    case Text text when separated:
                        result.Append(text.Text);
                        formatted ??= run;
                        break;

                    case DeletedText text when separated && revisions:
                        result.Append(text.Text);
                        formatted ??= run;
                        break;

                    case TabChar when separated:
                        result.Append('\t');
                        formatted ??= run;
                        break;

                    default:
                        return null;
                }
            }
        }

        return null;
    }

    /// <summary><c>w:fldSimple</c>: o mesmo nó, que volta na forma complexa, a mesma para o Word.</summary>
    private Node? ReadSimpleField(SimpleField field, RunProperties inherited, string? hyperlink)
    {
        // Pelo mesmo motivo de ReadField.
        if (field.FieldLock is not null || field.Dirty is not null || field.GetFirstChild<FieldData>() is not null)
        {
            return null;
        }

        var result = new System.Text.StringBuilder();
        Run? formatted = null;

        foreach (var child in field.ChildElements)
        {
            if (child is not Run run) return null;
            foreach (var piece in run.ChildElements)
            {
                switch (piece)
                {
                    case RunProperties: break;
                    case Text text: result.Append(text.Text); break;
                    case TabChar: result.Append('\t'); break;
                    default: return null;
                }
            }

            formatted ??= run;
        }

        var node = Node.Of("field")
            .With("instr", field.Instruction?.Value ?? string.Empty)
            .With("result", result.ToString());
        node.Marks = formatted is null ? null : MarksOfRun(formatted, inherited, hyperlink);
        if (result.Length > 0) _paragraphHasContent = true;
        return node;
    }

    private IEnumerable<Node> ReadRun(Run run, RunProperties inherited, string? hyperlink)
    {
        var marks = MarksOfRun(run, inherited, hyperlink);

        Node? customMark = null;

        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case Text text when customMark is not null:
                    customMark.With("mark", text.Text);
                    customMark = null;
                    break;

                // O número é a ordem no documento; o corpo da nota vai dentro do nó.
                case FootnoteReference or EndnoteReference when notes && _textBoxDepth == 0:
                {
                    var noteRef = ReadNoteRef(element);
                    if (IsCustomMark(element)) customMark = noteRef;
                    _paragraphHasContent = true;
                    yield return noteRef;
                    break;
                }

                // O Word desenha o número do corpo, e a gravação o refaz (NotesWriter).
                case FootnoteReferenceMark or EndnoteReferenceMark when _oidPrefix is not null:
                    break;

                case Text text:
                    if (text.Text.Length > 0)
                    {
                        _paragraphHasContent = true;
                        yield return new Node { Type = "text", Text = text.Text, Marks = marks };
                    }

                    break;

                case DeletedText deleted when revisions:
                    if (deleted.Text.Length > 0)
                    {
                        _paragraphHasContent = true;
                        yield return new Node { Type = "text", Text = deleted.Text, Marks = marks };
                    }

                    break;

                case TabChar:
                    yield return new Node { Type = "text", Text = "\t", Marks = marks };
                    break;

                case Break br:
                    // A quebra de coluna vira propriedade do bloco, como a de página;
                    // no rascunho de antes das seções era quebra de linha.
                    if (br.Type is not null && br.Type.Value == BreakValues.Page)
                    {
                        yield return Node.Of("pageBreak");
                    }
                    else if (br.Type is not null && br.Type.Value == BreakValues.Column && sections)
                    {
                        yield return Node.Of("columnBreak");
                    }
                    else
                    {
                        var hardBreak = Node.Of("hardBreak");
                        hardBreak.Marks = RevisionMarks();
                        yield return hardBreak;
                    }

                    break;

                case DocumentFormat.OpenXml.Wordprocessing.Drawing:
                    foreach (var node in ReadShape(element)) yield return node;
                    break;

                case Picture picture:
                    // VML: um documento do Word 2003 traz imagens assim no corpo.
                    foreach (var node in ReadShape(picture)) yield return node;
                    break;

                case AlternateContent alternate:
                    // O `mc:Choice` e o `mc:Fallback` são o mesmo conteúdo: lê-se um
                    // ramo só, senão cada caixa apareceria em dobro.
                    var branch = (OpenXmlElement?)alternate.GetFirstChild<AlternateContentChoice>()
                                 ?? alternate.GetFirstChild<AlternateContentFallback>();
                    if (branch is not null)
                    {
                        foreach (var node in ReadShape(branch)) yield return node;
                    }

                    break;

                case RunProperties:
                case LastRenderedPageBreak:
                    break;

                // A referência que divide o run com texto: rara, e a gravação a declara.
                case CommentReference when comments:
                    break;

                case FieldChar:
                case FieldCode:
                    // O campo que o nó não representa: o resultado em cache vem no run
                    // seguinte, e reescrito viraria texto comum.
                    inventory.NoteInvisible(Inventory.Fields);
                    _unrepresentedField = true;
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }
    }


    /// <summary>
    /// A imagem e o texto de dentro da forma. Uma caixa de texto é conteúdo: a capa
    /// do modelo de manual guarda o título numa. Sem âncora, o texto entra na linha
    /// em que a forma está, a melhor aproximação sem paginar.
    /// </summary>
    private IEnumerable<Node> ReadShape(OpenXmlElement shape)
    {
        // Ancorado sai do fluxo e vira propriedade do parágrafo âncora, como no
        // Word; `wp:inline` fica no fluxo. Menos o ancorado que o LibreOffice põe
        // onde o fluxo já o poria — ver AnchorReader.FlowsWithText.
        if (AnchorReader.AnchorOf(shape) is { } anchor && !AnchorReader.FlowsWithText(anchor))
        {
            foreach (var floating in DescribeAnchored(shape, anchor)) _paragraphFloats.Add(floating);
            yield break;
        }

        if (ReadImage(shape) is { } image)
        {
            _paragraphHasContent = true;
            yield return image;
        }

        // Caixa sem âncora não tem posição própria: o texto entra na linha.
        foreach (var node in ReadTextBoxesInline(shape)) yield return node;
    }

    /// <summary>Uma imagem, uma caixa de texto, ou nada.</summary>
    private IEnumerable<FloatDto> DescribeAnchored(
        OpenXmlElement shape,
        Drawing.Wordprocessing.Anchor anchor)
    {
        if (ImageSourceOf(shape) is { } src)
        {
            yield return AnchorReader.Describe(anchor, "image", src, null);
            yield break;
        }

        foreach (var box in OutermostTextBoxes(shape))
        {
            // Só o que não se desenha entra no inventário.
            var look = ShapeLook.Of(box);
            if (!look.Complete) inventory.NoteInvisible(Inventory.Shapes);

            var content = new List<Node>();
            foreach (var paragraph in box.Descendants<Paragraph>())
            {
                if (paragraph.Ancestors<TextBoxContent>().First() != box) continue;
                content.Add(ReadParagraphOf(paragraph));
            }

            if (content.Count > 0)
            {
                yield return AnchorReader.Describe(anchor, "text", null, content) with
                {
                    Fill = look.Fill,
                    Line = look.Line,
                    LineWidthPt = look.LineWidthPt,
                    Dash = look.Dashed,
                };
            }
        }
    }

    /// <summary>A caixa é um fluxo próprio: os parágrafos dela são parágrafos, e não linhas emendadas.</summary>
    private Node ReadParagraphOf(Paragraph paragraph)
    {
        var (effective, inheritedRun) = _styles.Resolve(paragraph.ParagraphProperties);

        var node = Node.Of("paragraph");
        if (AlignmentOf(effective) is { } alignment) node.With("textAlign", alignment);

        var outer = _directRuns;
        _directRuns = false;
        _textBoxDepth++;
        var content = ReadInline(paragraph, inheritedRun);
        _textBoxDepth--;
        _directRuns = outer;
        if (content.Count > 0) node.Content = content;
        return node;
    }

    /// <summary>Sem âncora, o conteúdo entra na linha.</summary>
    private IEnumerable<Node> ReadTextBoxesInline(OpenXmlElement shape)
    {
        foreach (var box in OutermostTextBoxes(shape))
        {
            // A caixa não é desenhada: a moldura dela se perde de vista, e o aviso fala disso.
            var look = ShapeLook.Of(box);
            if (!look.Complete || look.Draws) inventory.NoteInvisible(Inventory.Shapes);

            foreach (var paragraph in box.Descendants<Paragraph>())
            {
                if (paragraph.Ancestors<TextBoxContent>().First() != box) continue;

                var (_, inheritedRun) = _styles.Resolve(paragraph.ParagraphProperties);
                var afterSomething = _paragraphHasContent;

                var outer = _directRuns;
                _directRuns = false;
                _textBoxDepth++;
                var inline = ReadInline(paragraph, inheritedRun);
                _textBoxDepth--;
                _directRuns = outer;
                if (inline.Count == 0) continue;

                if (afterSomething) yield return Node.Of("hardBreak");
                _paragraphHasContent = true;

                foreach (var node in inline) yield return node;
            }
        }
    }

    /// <summary>
    /// Para na primeira caixa de cada ramo: a de dentro é alcançada pela recursão
    /// de <see cref="ReadInline"/>, e as duas rotas escreveriam o texto duas vezes.
    /// </summary>
    private static IEnumerable<TextBoxContent> OutermostTextBoxes(OpenXmlElement root) =>
        TextBoxNav.Outermost(root);

    /// <summary>Separado de <see cref="ReadImage"/>: o objeto ancorado precisa dos bytes sem o nó.</summary>
    private string? ImageSourceOf(OpenXmlElement drawing)
    {
        var blip = drawing.Descendants<Drawing.Blip>().FirstOrDefault();
        var relationshipId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;

        if (!_owner.TryGetPartById(relationshipId, out var found) || found is not ImagePart image)
        {
            inventory.NoteLoss("imagem em formato não suportado");
            return null;
        }

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);

        return $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}";
    }

    /// <summary>
    /// A captura ancorada ao topo do parágrafo vai para o começo dele, como o Word e
    /// o LibreOffice a desenham; ancorada à linha, fica onde está.
    /// </summary>
    private void TopAnchoredFirst(List<Node> content)
    {
        if (_topAnchored.Count == 0) return;
        var moved = content.Where(_topAnchored.Contains).ToList();
        if (moved.Count == 0 || content.Take(moved.Count).SequenceEqual(moved)) return;
        content.RemoveAll(_topAnchored.Contains);
        content.InsertRange(0, moved);
    }

    /// <summary>
    /// A imagem em linha (<c>wp:inline</c>) e a ancorada onde o fluxo já a poria —
    /// ver <see cref="AnchorReader.FlowsWithText"/>.
    /// </summary>
    private Node? ReadImage(OpenXmlElement drawing)
    {
        if (ImageSourceOf(drawing) is not { } src) return null;

        var node = Node.Of("image").With("src", src);

        // Ancorada ao parágrafo, ainda que no lugar do fluxo: o parágrafo que ancora
        // ocupa a linha dele além da imagem (11,55 pt medidos no LibreOffice).
        if (AnchorReader.AnchorOf(drawing) is { } anchor)
        {
            node.With("anchored", true);
            var vertical = anchor.GetFirstChild<Drawing.Wordprocessing.VerticalPosition>();
            var from = vertical?.RelativeFrom?.Value;
            // Deslocamento até 1 pt (635 EMU no corpus) é o zero que o LibreOffice grava.
            var offset = long.TryParse(vertical?.PositionOffset?.Text, out var emus) ? Math.Abs(emus) : 0;
            if ((from is null || from == Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph) &&
                offset <= 12700)
            {
                _topAnchored.Add(node);
            }
        }

        // O texto alternativo (`wp:docPr/@descr`) vira o `alt`; `descr=""` é o padrão
        // de quem nunca o preencheu, e não vira atributo.
        var properties = drawing.Descendants<Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
        if (properties?.Description?.Value is { Length: > 0 } description)
        {
            node.With("alt", description);
        }

        var extent = drawing.Descendants<Drawing.Wordprocessing.Extent>().FirstOrDefault();
        if (extent?.Cx?.Value is { } wide && extent.Cy?.Value is { } tall && wide > 0 && tall > 0)
        {
            // `wp:extent` mede antes de girar: num quarto de volta, largura e altura trocam.
            var (across, down) = IsQuarterTurned(drawing) ? (tall, wide) : (wide, tall);

            // As duas medidas: sem a altura, a paginação mediria a folha sem a
            // imagem, e a proporção seria a do arquivo. EMU → px: 914400 por
            // polegada, 96 px por polegada.
            node.With("width", (int)Math.Round(across * 96.0 / 914400));
            node.With("height", (int)Math.Round(down * 96.0 / 914400));
        }

        return node;
    }

    /// <summary>
    /// <c>a:rot</c> vem em 60000 avos de grau, e pode ser negativo. Só o quarto de
    /// volta troca largura e altura.
    /// </summary>
    internal static bool IsQuarterTurned(OpenXmlElement drawing)
    {
        var rotation = drawing.Descendants<Drawing.Transform2D>().FirstOrDefault()?.Rotation?.Value;
        if (rotation is null) return false;

        var degrees = ((rotation.Value / 60000.0) % 360 + 360) % 360;
        return degrees is (> 45 and < 135) or (> 225 and < 315);
    }

    private string? HyperlinkTargetOf(Hyperlink link)
    {
        var id = link.Id?.Value;

        // O link interno tem `w:anchor`, e não relacionamento: no editor vira `#nome`.
        if (string.IsNullOrEmpty(id))
        {
            return references && link.Anchor?.Value is { Length: > 0 } anchor ? "#" + anchor : null;
        }

        try
        {
            return _owner.HyperlinkRelationships
                .FirstOrDefault(relationship => relationship.Id == id)?.Uri.ToString();
        }
        catch (UriFormatException)
        {
            return null;
        }
    }


    private Node ReadTable(Table table)
    {
        var rows = new List<Node>();

        // A grade da tabela, que o Word usa e o editor espelha no `colgroup`.
        var grid = TableLook.GridWidths(table);

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = new List<Node>();

            // `w:tblHeader` é a linha de cabeçalho do editor; presente sem `w:val` já é "sim".
            var repeat = row.TableRowProperties?.GetFirstChild<TableHeader>();
            var header = repeat is not null
                         && !(repeat.Val is { } declared && declared.Value == OnOffOnlyValues.Off);

            // `w:gridBefore`: a linha pode começar colunas adiante.
            var column = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value is > 0 and var before
                ? before
                : 0;

            foreach (var cell in row.Elements<TableCell>())
            {
                // A tabela aninhada é um bloco, e o `tableCell` do editor aceita bloco.
                var contents = new List<Node>();
                foreach (var child in cell.ChildElements)
                {
                    switch (child)
                    {
                        case Paragraph paragraph: contents.Add(ReadParagraph(paragraph, flat: true)); break;
                        case Table nested: contents.Add(ReadTable(nested)); break;
                        case TableCellProperties: break;
                        default: inventory.NoteInvisibleElement(child.LocalName); break;
                    }
                }

                if (contents.Count == 0) contents.Add(Node.Of("paragraph"));

                var node = Node.Of(header ? "tableHeader" : "tableCell");
                node.Content = contents;

                // **Sempre** escritos: o editor devolve `colspan` e `rowspan` em toda
                // célula, e omiti-los faria toda tabela ser regenerada.
                var span = cell.TableCellProperties?.GridSpan?.Val?.Value;
                var spanned = span is > 1 ? span.Value : 1;
                node.With("colspan", spanned);

                // Mesclagem vertical fica no `w:vMerge` do arquivo.
                node.With("rowspan", 1);

                // Uma medida por coluna, em pixels do CSS, como o TableKit usa.
                if (grid.Count >= column + spanned)
                {
                    var widths = new JsonArray();
                    for (var index = 0; index < spanned; index++) widths.Add(grid[column + index]);
                    node.With("colwidth", widths);
                }

                column += spanned;

                // Só quando o arquivo os declara: ausente e nulo são a mesma afirmação.
                if (TableLook.Shading(cell.TableCellProperties) is { } fill) node.With("shading", fill);
                if (TableLook.Borders(cell.TableCellProperties) is { } borders) node.With("borders", borders);

                cells.Add(node);
            }

            var rowNode = Node.Of("tableRow");
            rowNode.Content = cells;

            if (revisions && Revisions.BlockRevisionOf(row.TableRowProperties) is { } rowRevision)
            {
                rowNode.With("rowRevision", rowRevision);
            }

            rows.Add(rowNode);
        }

        var tableNode = Node.Of("table");
        tableNode.Content = rows;

        // Para a tela medir a linha como o papel; o `w:tblPr` volta como estava.
        tableNode.With("cellMargins", string.Join(' ', TableLook.CellMargins(table, part)));
        return tableNode;
    }

}

using System.Text.Json;
using System.Text.Json.Nodes;
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
public sealed class BodyReader(MainDocumentPart part, Inventory inventory, bool flatten = false)
{
    /// <summary>Passo de recuo do Word: meia polegada.</summary>

    private readonly NumberingReader _numbering = new(part);
    private readonly StyleResolver _styles = new(part);
    private readonly FontTable _fonts = new(part);
    private int _nextId = 1;

    /// <summary>
    /// Já saiu conteúdo no parágrafo que está sendo lido?
    /// </summary>
    /// <remarks>
    /// Existe por uma razão só: decidir se uma caixa de texto abre linha nova.
    /// As caixas de um mesmo parágrafo estão em `w:r` diferentes, então nenhuma
    /// delas consegue ver o que a anterior escreveu — e a resposta precisa
    /// atravessar essa fronteira.
    /// </remarks>
    private bool _paragraphHasContent;

    /// <summary>
    /// Os trechos do parágrafo em leitura levam só o que difere do estilo?
    /// </summary>
    /// <remarks>
    /// Só no parágrafo que o CSS dos estilos desenha — o solto no corpo, lido sem
    /// achatar. O de lista, de célula, de caixa de texto e o do rascunho antigo
    /// continuam com a formatação inteira em cada trecho: ali nenhuma regra de
    /// estilo chega para preencher o que faltasse.
    /// </remarks>
    private bool _directRuns;

    /// <summary>
    /// Objetos ancorados encontrados no parágrafo que está sendo lido.
    /// </summary>
    /// <remarks>
    /// São descobertos no meio da leitura de linha, mas pertencem ao **bloco**:
    /// não ocupam lugar no fluxo, e a posição deles é dada em relação à página,
    /// à margem ou ao próprio parágrafo. Acumulam aqui e são anexados ao nó do
    /// parágrafo no fim, que é onde quem desenha vai procurá-los.
    /// </remarks>
    private readonly List<FloatDto> _paragraphFloats = [];

    /// <summary>
    /// As capturas ancoradas ao **topo do parágrafo** lidas no parágrafo corrente
    /// — ver <see cref="TopAnchoredFirst"/>.
    /// </summary>
    private readonly HashSet<Node> _topAnchored = new(ReferenceEqualityComparer.Instance);

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
        var openLists = new List<(Node List, string Kind, int Level, int NumId)>();

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case Paragraph paragraph:
                {
                    // Só o parágrafo solto no corpo deixa de ser achatado: é ele
                    // que as regras dos estilos alcançam (`.page__content > p`).
                    // O item de lista e o parágrafo de célula têm regras próprias
                    // em `content-styles.ts`, e continuam levando o efetivo.
                    var numbered = _numbering.ListKindOf(paragraph.ParagraphProperties);
                    var node = ReadParagraph(paragraph, flat: flatten || numbered is not null);
                    var list = node.Type == "pageBreak" ? null : numbered;

                    if (list is null)
                    {
                        openLists.Clear();
                        blocks.Add(NewBlock(element, node));
                        content.Add(node);
                        break;
                    }

                    var (kind, level) = (list.Kind, list.Level);
                    while (openLists.Count > 0 && openLists[^1].Level > level)
                    {
                        openLists.RemoveAt(openLists.Count - 1);
                    }

                    // No mesmo nível, outra numeração é outra lista — ainda que do
                    // mesmo tipo. Juntadas, a segunda passava a ser contada pela
                    // primeira, e a tela numerava diferente do Word.
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

                        // A marca e o recuo são do nível, e não do parágrafo: é
                        // a lista que os desenha. Sem eles a bolinha do CSS
                        // aparece no lugar do quadrado que o documento pede, e
                        // o item sai colado na margem.
                        // O `numId` viaja no nó da lista porque é o que a
                        // gravação precisa para continuar apontando a **mesma**
                        // numeração do arquivo. Sem ele o escritor gravava
                        // `w:numId w:val="0"` — que no formato quer dizer "sem
                        // numeração" — e a lista voltava como parágrafos comuns.
                        listNode.With("numId", list.NumberingId);

                        if (list.Marker is { } marker) listNode.With("marker", marker);

                        // O `start` que o editor materializa em toda lista numerada.
                        // Não conta nada — quem conta é a definição —, mas entra na
                        // impressão digital do item de fora quando a lista é
                        // aninhada, e ausente de um lado só o item era reescrito.
                        if (kind == "orderedList") listNode.With("start", 1);
                        if (list.IndentMm is { } indent) listNode.With("indentMm", indent);
                        if (list.HangingMm is { } hanging) listNode.With("hangingMm", hanging);

                        // A definição inteira só onde a numeração muda: a sublista
                        // da mesma numeração a encontra na lista de fora, e repeti-la
                        // em cada nível só engordaria o documento.
                        if (depth == 0 || parent.NumId != list.NumberingId)
                        {
                            listNode.With("numbering", list.Definition);
                        }

                        // O nível do arquivo, quando a árvore não o diz sozinha: a
                        // lista que começa no nível 2, sem 0 e 1 antes, mora no
                        // topo da árvore — e sem isto voltava ao arquivo no nível 0,
                        // com outra marca e outra conta.
                        if (level != (depth == 0 ? 0 : parent.Level + 1)) listNode.With("level", level);

                        if (depth > 0)
                        {
                            // Lista aninhada mora dentro do último item da de fora.
                            var parentItems = parent.List.Content!;
                            (parentItems[^1].Content ??= []).Add(listNode);
                        }
                        else
                        {
                            content.Add(listNode);
                        }

                        openLists.Add((listNode, kind, level, list.NumberingId));
                    }

                    // O id fica no `listItem`, não no parágrafo: é o item que
                    // corresponde a um `w:p` do arquivo.
                    var item = Node.Of("listItem");
                    item.Content = [node];
                    openLists[^1].List.Content!.Add(item);
                    CarrySpacing(openLists[^1].List, node);
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

    /// <summary>
    /// A lista herda o espaçamento dos parágrafos das pontas.
    /// </summary>
    /// <remarks>
    /// No arquivo a lista não existe como bloco: o que existe são parágrafos com
    /// numeração, cada um com o seu espaçamento. Na árvore do editor a lista é
    /// um elemento de verdade, e um elemento sem espaçamento declarado recebe o
    /// do editor — o mesmo `0.6em` que o parágrafo importado já não recebe. Num
    /// documento com seis listas isso somava quinze milímetros de ar que o Word
    /// não tem, o bastante para empurrar a última imagem para uma folha nova.
    ///
    /// Antes do primeiro item vale o espaço de antes dele; depois do último,
    /// o de depois — que é o que o Word desenha.
    /// </remarks>
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
        var oid = "b" + _nextId++;
        extracted.With("oid", oid);
        return new Block(oid, source, extracted);
    }

    // --- parágrafos ---------------------------------------------------------

    private Node ReadParagraph(Paragraph paragraph, bool flat)
    {
        var direct = paragraph.ParagraphProperties;

        // Formatação efetiva: padrões do documento, estilo e formatação direta.
        // Ler só a direta é o que fazia um documento cheio de estilos abrir
        // praticamente sem formatação.
        var (effective, inheritedRun) = _styles.Resolve(direct);

        _paragraphHasContent = false;
        _paragraphFloats.Clear();
        _topAnchored.Clear();
        var outer = _directRuns;
        _directRuns = !flat;
        var content = ReadInline(paragraph, inheritedRun);
        _directRuns = outer;

        // Uma quebra de página sozinha no parágrafo é o nó `pageBreak`, não um
        // parágrafo vazio com uma quebra dentro.
        //
        // **Menos** quando o parágrafo ancora alguma coisa. O nó de quebra mora
        // no vão entre duas folhas, e um objeto ancorado nele cai na folha de
        // baixo; o parágrafo do arquivo está na de cima, antes da quebra. Foi
        // assim que a marca vertical da capa do modelo de manual apareceu no
        // topo da segunda folha em vez de correr pela lateral da primeira,
        // enquanto o LibreOffice a desenhava na capa.
        //
        // Preservado o parágrafo, a quebra vira propriedade dele — o mesmo
        // caminho da quebra no meio do texto, logo abaixo.
        if (content.Count == 1 && content[0].Type == "pageBreak" && _paragraphFloats.Count == 0)
        {
            return content[0];
        }

        // Quebra **no meio** de um parágrafo: `w:br w:type="page"` dentro de um
        // `w:r`, que é como o Word grava "a partir daqui é outra página" sem
        // fechar o parágrafo.
        //
        // Emiti-la ali dentro punha um nó de bloco em posição de linha — inválido
        // no editor, e caro de um jeito difícil de rastrear: serializado, um
        // `<div>` dentro de `<p>` faz o analisador de HTML fechar o parágrafo e
        // desalojar o `div`, e um documento de 15 blocos vira 17 elementos. Os
        // índices deixam de casar e o papel corta em lugar diferente da tela.
        //
        // Vira uma propriedade do bloco: a folha termina **depois** deste
        // parágrafo. É aproximação quando ainda há texto depois da quebra dentro
        // do mesmo parágrafo — esse texto desce junto em vez de abrir a página —
        // e é exato no caso comum, que é a quebra encerrando o parágrafo.
        var breakAfter = content.RemoveAll(child => child.Type == "pageBreak") > 0;
        TopAnchoredFirst(content);

        var node = (HeadingLevelOf(direct) ?? _styles.HeadingLevelByName(direct?.ParagraphStyleId?.Val?.Value)) is { } level
            ? Node.Of("heading").With("level", level)
            : Node.Of("paragraph");

        if (breakAfter) node.With("breakAfter", true);

        WithFloats(node);

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

        if (flat) Flattened(node, alignment, effective, direct, inheritedRun);
        else DirectOnly(node, viaTabs, effective, direct, inheritedRun);

        // O parágrafo que só carrega a marca de seção não é uma linha de texto.
        //
        // No OOXML a seção termina num `w:sectPr` guardado dentro do `w:pPr` de
        // um parágrafo vazio: o parágrafo **é** a marca. O LibreOffice não lhe
        // dá altura nenhuma, e é ele quem grava documentos assim — o de
        // evidências do corpus tem sete seções, todas com a mesma geometria, e
        // seis marcas espalhadas pelo meio do texto. Cada uma valia uma linha
        // aqui, e o texto ia descendo folha após folha.
        //
        // A marca continua no modelo, e não é descartada: é ela que a gravação
        // devolve ao arquivo, e sem ela as seções do documento sumiriam.
        if (direct?.SectionProperties is not null && content.Count == 0)
        {
            node.With("sectionMark", true);
        }

        node.Content = content.Count == 0 ? null : content;
        return node;
    }

    /// <summary>
    /// A formatação **efetiva** no bloco: padrões, estilo e direta, achatados.
    /// </summary>
    /// <remarks>
    /// Continua valendo onde as regras dos estilos não chegam — item de lista e
    /// célula — e no modo <c>flatten</c>, que é a leitura de referência de um
    /// rascunho gravado antes de o bloco carregar só a formatação direta: ali os
    /// nós vieram achatados, e compará-los com uma leitura que não achata faria
    /// todo bloco parecer mudado.
    /// </remarks>
    private void Flattened(
        Node node,
        string? alignment,
        ParagraphProperties effective,
        ParagraphProperties? direct,
        RunProperties inheritedRun)
    {
        if (alignment is not null) node.With("textAlign", alignment);


        // Zero, sempre. O nível é do editor — `Ctrl+]` trabalha em passos, e um
        // passo vale 2,5em, que a 10 pt são 25 pt e não os 36 pt que 720 twips
        // pedem. O recuo do arquivo vem logo abaixo, na medida em que ele o
        // declara; escrever os dois somava um recuo que ninguém pediu.
        //
        // Continua sendo escrito porque o editor declara `indent` com padrão 0 e
        // devolve o atributo em todo parágrafo: omiti-lo aqui fazia os dois
        // lados descreverem o mesmo bloco de formas diferentes, e a comparação
        // que decide o que preservar na gravação dizia "mudou" em bloco que
        // ninguém tocou.
        node.With("indent", 0);

        // O recuo **em milímetros**, que é como o arquivo o declara. Enquanto
        // era só o nível, todo parágrafo recuado saía 30% mais estreito do que
        // no LibreOffice, e a captura dentro dele encolhia junto: era o que
        // deixava a legenda caber na folha em que o LibreOffice já não a punha.
        Measure(node, "indentMm", effective.Indentation?.Left?.Value);
        Measure(node, "indentRightMm", effective.Indentation?.Right?.Value);

        // A primeira linha, que anda para os dois lados: `w:firstLine` a empurra
        // e `w:hanging` a puxa. É o mesmo `text-indent` do CSS, e é o que faz o
        // parágrafo pendurado ter a primeira linha fora do recuo das demais.
        var firstLine = TwipsToMm(effective.Indentation?.FirstLine?.Value);
        var hanging = TwipsToMm(effective.Indentation?.Hanging?.Value);
        if (firstLine is > 0) node.With("firstLineMm", firstLine.Value);
        else if (hanging is > 0) node.With("firstLineMm", -hanging.Value);

        // O fundo do parágrafo é o que transforma `Heading1` numa barra
        // colorida neste corpus — sem ele o título vira texto solto.
        if (ShadingOf(effective) is { } background) node.With("background", background);

        // Espaçamento e entrelinha são **sempre** escritos, como o recuo e pela
        // mesma razão: silêncio no arquivo não significa "use o seu padrão",
        // significa zero. Enquanto ficavam ausentes, o `margin-top: 0.6em` e o
        // `line-height: 1.5` do editor — que existem para o documento em branco
        // — reapareciam em cada parágrafo importado. Num documento de 48
        // parágrafos isso somava mais de uma página de ar que o Word não tem, e
        // era o que fazia a imagem seguinte descer para a folha de baixo.
        var spacing = effective.SpacingBetweenLines;
        node.With("spaceBefore", TwipsToPt(spacing?.Before?.Value) ?? 0);
        node.With("spaceAfter", TwipsToPt(spacing?.After?.Value) ?? 0);
        // A entrelinha depende da fonte com que a linha é medida, e é a marca
        // do parágrafo que a diz — a mesma que dá altura ao parágrafo vazio.
        var markFont = _styles.ResolveMark(inheritedRun, direct).RunFonts?.Ascii?.Value;
        node.With("lineHeight", LineHeightOf(spacing, LineMetrics.Of(markFont)));

        // A fonte **do bloco**, e não só a dos runs. A altura da linha nasce da
        // fonte do próprio elemento: sem isto, um parágrafo de 10 pt dentro de
        // um bloco que o CSS declara com 12 pt continua ocupando 12 pt de
        // altura — e um `Heading1` de 10 pt vira uma barra alta demais, porque
        // o editor desenha títulos em 22 pt.
        //
        // Vem da **marca de parágrafo** (`w:pPr/w:rPr`), e não do estilo só: é
        // ela que o Word usa para medir a linha e para dar altura ao parágrafo
        // vazio. Sem ela, um parágrafo vazio de Verdana 10 pt ocupava os 12 pt
        // do padrão do editor — meia linha a mais, vinte vezes no documento.
        var mark = _styles.ResolveMark(inheritedRun, direct);
        if (FontOf(mark) is { } font) node.With("fontFamily", font);
        if (FontSizeOf(mark) is { } size) node.With("fontSize", size);

        // "Manter com o próximo": o parágrafo não fica sozinho no pé da página.
        // É o que faz um rótulo descer junto com a imagem que ele apresenta —
        // e sem ler isto a quebra estimada cai um bloco depois da real.
        if (RunReader.IsOn(effective.KeepNext)) node.With("keepNext", true);

        // "Manter linhas juntas": a paginação corta parágrafos entre linhas, e
        // este é o parágrafo que pediu para não ser cortado.
        if (RunReader.IsOn(effective.KeepLines)) node.With("keepLines", true);

        // Viúvas e órfãs: ligado quando o arquivo cala, como o Word faz — então
        // só o desligado vale ser dito.
        if (effective.WidowControl is { } widow && !RunReader.IsOn(widow)) node.With("widowControl", false);
    }

    /// <summary>
    /// Só a formatação **direta** no bloco; o herdado vem do CSS dos estilos.
    /// </summary>
    /// <remarks>
    /// A regra é uma só, campo a campo: o arquivo declara a propriedade no
    /// `w:pPr` do parágrafo (ou na marca, `w:pPr/w:rPr`, para a fonte)? Então o
    /// bloco leva o valor **efetivo** dela, calculado como no achatamento; se
    /// cala, o bloco cala, e quem desenha é `style-css.ts`. Levar o efetivo, e
    /// não o atributo cru, é o que mantém a tela igual à de antes: o `w:ind` e o
    /// `w:spacing` se fundem atributo a atributo com os do estilo, e o valor que
    /// sai da fusão é o que o achatamento mostrava.
    ///
    /// Zero declarado **é** declaração: um `w:ind w:left="0"` desfaz o recuo do
    /// estilo, e omiti-lo traria o recuo de volta pela regra do estilo.
    ///
    /// Duas exceções derivam de outra coisa que não o `w:pPr`, e por isso podem
    /// aparecer sem declaração direta: o alinhamento por tabulação
    /// (<see cref="TabAlignmentOf"/>), que nasce do conteúdo; e a entrelinha
    /// quando a fonte da marca é direta — o múltiplo é medido sobre a altura
    /// natural da fonte, e a regra do estilo o mede com a fonte do estilo.
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

        // O nível de `Ctrl+]`, zero pelo mesmo motivo do achatamento: o editor
        // declara `indent` com padrão 0 e o devolve em todo parágrafo.
        node.With("indent", 0);

        // Recuo negativo sai como zero, que é o que o achatamento mostrava: o
        // bloco não desenha recuo para fora da margem.
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
            // `w:shd` sem cor sobre um estilo com fundo apaga o fundo do estilo:
            // sem dizê-lo, a regra do estilo o pintaria de volta.
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

        // Ligado ou desligado, se o parágrafo diz: `w:keepNext w:val="0"` existe
        // para desfazer o do estilo.
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
    /// Entrelinha, já em CSS. `w:line` com regra `auto` vem em 240-avos: 271
    /// significa 1,13 vez a altura natural da linha.
    /// </summary>
    /// <remarks>
    /// **Vez a altura natural**, e não vez o tamanho da fonte: é a diferença
    /// entre 12,98 pt e 11,3 pt numa linha de Arial 10 pt, e é o que fazia um
    /// documento caber em menos folhas aqui do que no LibreOffice.
    ///
    /// Por isso sai número sempre que se sabe qual arquivo de fonte o navegador
    /// vai usar, inclusive no espaçamento simples — que seria `normal` em CSS,
    /// mas cujo cálculo o Chromium arredonda para pixel inteiro: 15 px onde o
    /// LibreOffice usa 15,33. São 2 % por linha, o bastante para um documento
    /// de quinze folhas fechar em dezesseis.
    ///
    /// Quando a fonte não é uma das que o instalador leva, a substituta depende
    /// da máquina e não há altura honesta a declarar: fica `normal`, e quem
    /// mede é o navegador.
    ///
    /// `exact` e `atLeast` dizem a altura em twips, e viram pontos. Antes daqui
    /// as duas voltavam nulas, e um parágrafo com entrelinha travada em 9 pt
    /// era desenhado com a do editor.
    /// </remarks>
    /// <param name="decimals">
    /// Casas do múltiplo antes de multiplicar. Duas no bloco achatado, como
    /// sempre foi — mudar faria o rascunho antigo parecer editado. Quatro no
    /// bloco que só leva o direto: é a grade de 240-avos em que o múltiplo volta
    /// ao arquivo, e a mesma com que <see cref="StyleReader"/> entrega o do
    /// estilo — sem isso a entrelinha direta e a herdada seriam medidas em
    /// grades diferentes.
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

            // Fora dessa faixa é lixo do arquivo, e não pedido de espaçamento.
            if (factor is > 0.5 and < 4) return Multiple(factor, natural);
        }

        return Multiple(1, natural);
    }

    /// <summary>
    /// O múltiplo já resolvido na altura da fonte.
    /// </summary>
    /// <remarks>
    /// Fonte que o instalador não leva cai numa substituta que depende da
    /// máquina. Sem múltiplo declarado, quem mede melhor é o navegador, e o
    /// leitor sai da frente com `normal`. Com múltiplo declarado não dá para
    /// sair da frente — aplicá-lo sobre o tamanho da fonte erra por 15 % —, e
    /// então vale o palpite de 1,15: é a altura de quase toda fonte latina, e a
    /// das substitutas que o LibreOffice escolhe.
    /// </remarks>
    private static string Multiple(double factor, double? natural)
    {
        if (natural is not null) return Text(Math.Round(factor * natural.Value, 4));
        return factor == 1 ? "normal" : Text(Math.Round(factor * 1.1499, 4));
    }

    private static string Text(double value) =>
        value.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>
    /// Nível do título a partir do estilo do parágrafo.
    /// </summary>
    /// <remarks>
    /// O corpus usa `Heading1`; o Word também grava `heading 1` e o LibreOffice
    /// `Ttulo1` (sem acento, porque o id do estilo não os aceita). Aceitar as
    /// três formas custa uma linha e evita que todo título vire texto normal.
    /// </remarks>
    private static int? HeadingLevelOf(ParagraphProperties? properties) =>
        HeadingLevelOfStyle(properties?.ParagraphStyleId?.Val?.Value);

    /// <summary>
    /// O mesmo, a partir do identificador de estilo cru.
    /// </summary>
    /// <remarks>
    /// Visível para <see cref="ParagraphFormat"/>: quem grava precisa saber se o
    /// estilo que o modelo carrega já é um estilo de título, para não trocar o
    /// `Ttulo1` do documento por um `Heading1` que ele não define.
    /// </remarks>
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

    /// <summary>Anexa ao bloco os objetos ancorados que o parágrafo trouxe.</summary>
    private Node WithFloats(Node node)
    {
        if (_paragraphFloats.Count > 0)
        {
            node.With("floats", JsonSerializer.SerializeToNode(_paragraphFloats, DocxJson.Options));
        }

        return node;
    }

    /// <summary>Escreve a medida no nó quando ela existe e não é zero.</summary>
    private static void Measure(Node node, string name, string? twips)
    {
        if (TwipsToMm(twips) is { } value and > 0) node.With(name, value);
    }

    /// <summary>1 twip = 1/1440 de polegada.</summary>
    private static double? TwipsToMm(string? twips) =>
        int.TryParse(twips, out var value) ? Math.Round(value * 25.4 / 1440, 2) : null;

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
                    inventory.NoteInvisible(Inventory.Comments);
                    break;

                case InsertedRun inserted:
                    inventory.NoteInvisible(Inventory.TrackedChanges);
                    nodes.AddRange(ReadInline(inserted, inherited, hyperlink));
                    break;

                case DeletedRun:
                    // Texto marcado como excluído não deve aparecer na tela.
                    inventory.NoteInvisible(Inventory.TrackedChanges);
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
        // O herdado vai junto para que o "desligado" direto sobre um estilo que
        // liga vire marca — a tela desenha o estilo, e sem ela o trecho voltaria
        // negrito. O rascunho antigo não as conhece, e a leitura dele não as dá.
        var marks = RunReader.MarksOf(
            _styles.ResolveRun(inherited, run.RunProperties),
            hyperlink,
            _fonts,
            flatten ? null : inherited,
            directOnly: _directRuns);

        // O estilo de caractere do trecho (`w:rStyle`). A leitura não o resolve —
        // quem o desenha é o CSS dos estilos, como no parágrafo —, e a marca é o
        // que o faz voltar ao arquivo.
        if (!flatten && run.RunProperties?.RunStyle?.Val?.Value is { Length: > 0 } characterStyle)
        {
            (marks ??= []).Add(Mark.Of("charStyle", "styleId", characterStyle));
        }

        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case Text text:
                    if (text.Text.Length > 0)
                    {
                        _paragraphHasContent = true;
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
                    foreach (var node in ReadShape(element)) yield return node;
                    break;

                case Picture picture:
                    // VML antigo. No corpus só aparece nos cabeçalhos, mas um
                    // documento do Word 2003 traz imagens assim no corpo.
                    foreach (var node in ReadShape(picture)) yield return node;
                    break;

                case AlternateContent alternate:
                    // Word grava a mesma forma duas vezes: `mc:Choice` em
                    // DrawingML e `mc:Fallback` no VML que o Word 2007 entendia.
                    // São o mesmo conteúdo, então lê-se um ramo só — ler os dois
                    // faria cada caixa de texto aparecer em dobro na tela.
                    //
                    // Nada de inventário aqui: quem sabe se a forma tem moldura
                    // que não reproduzimos é quem a lê, mais abaixo. Este aviso
                    // era dado a toda forma, imagem inclusive, e por isso
                    // aparecia em todo documento que tivesse uma caixa.
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

                case FieldChar:
                case FieldCode:
                    // `PAGE` e afins: o valor só existe na paginação. O texto
                    // que o Word deixou em cache vem no run seguinte.
                    inventory.NoteInvisible(Inventory.Fields);
                    break;

                default:
                    inventory.NoteInvisibleElement(element.LocalName);
                    break;
            }
        }
    }

    // --- formas e caixas de texto -------------------------------------------

    /// <summary>
    /// Forma → o que dela cabe numa linha de texto: a imagem e o que estiver
    /// escrito dentro dela.
    /// </summary>
    /// <remarks>
    /// Uma caixa de texto é conteúdo, não decoração. Antes daqui ela era
    /// registrada no inventário e descartada, e um documento cujo título mora
    /// dentro de uma caixa — a capa do modelo de manual é assim — abria sem
    /// título nenhum. O aviso dizia "formas e caixas de texto", que descreve o
    /// que aconteceu sem dizer o que sumiu.
    ///
    /// O texto entra na linha onde a forma está ancorada, e não na posição da
    /// página em que o Word a desenha: não há layout flutuante aqui, e a
    /// escolha é entre o texto no lugar aproximado ou o texto em lugar nenhum.
    /// A âncora é a melhor aproximação que existe sem paginar.
    ///
    /// O inventário continua marcando a forma — a moldura, a posição e o
    /// preenchimento realmente não aparecem, e a gravação cirúrgica ainda perde
    /// a forma inteira se o parágrafo âncora for editado. É isso que mantém o
    /// documento em somente leitura.
    /// </remarks>
    private IEnumerable<Node> ReadShape(OpenXmlElement shape)
    {
        // Ancorado é objeto **fora do fluxo**: no Word ele não empurra o texto,
        // mora numa posição da folha e pode até ficar atrás dela. Lê-lo como
        // bloco no meio do texto punha a marca vertical da capa como uma faixa
        // deitada de página inteira, empurrando tudo para baixo — e a contagem
        // de páginas ia junto.
        //
        // Sai do fluxo e vira propriedade do parágrafo âncora. O que continua
        // aqui é o objeto **no fluxo** (`wp:inline`), que é imagem no meio da
        // linha e deve mesmo ocupar lugar.
        //
        // Nem todo ancorado tem posição de verdade. É assim que o LibreOffice
        // grava "imagem no próprio parágrafo": ancorada ao parágrafo, sem
        // deslocamento, centralizada na coluna. Tirar essas do fluxo encolheu um
        // documento de trinta capturas de doze folhas para quatro, com as
        // imagens empilhadas umas sobre as outras. Ver AnchorReader.FlowsWithText.
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

        // Caixa de texto sem âncora é rara e não tem posição própria: o texto
        // dela entra na linha, que é onde estaria de qualquer modo.
        foreach (var node in ReadTextBoxesInline(shape)) yield return node;
    }

    /// <summary>
    /// O que um desenho ancorado carrega: uma imagem, uma caixa de texto, ou nada.
    /// </summary>
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
            // A moldura e o preenchimento da caixa, quando dá para desenhá-los.
            // Só o que sobrar entra no inventário: um aviso que aparece em todo
            // documento com caixa de texto é um aviso que se aprende a ignorar.
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

    /// <summary>
    /// Um parágrafo de dentro de uma caixa, com a formatação dele resolvida.
    /// </summary>
    /// <remarks>
    /// A caixa é um fluxo de texto próprio, então os parágrafos dela são
    /// parágrafos de verdade — e não linhas emendadas, como eram quando o texto
    /// era despejado na linha do parágrafo âncora.
    /// </remarks>
    private Node ReadParagraphOf(Paragraph paragraph)
    {
        var (effective, inheritedRun) = _styles.Resolve(paragraph.ParagraphProperties);

        var node = Node.Of("paragraph");
        if (AlignmentOf(effective) is { } alignment) node.With("textAlign", alignment);

        var outer = _directRuns;
        _directRuns = false;
        var content = ReadInline(paragraph, inheritedRun);
        _directRuns = outer;
        if (content.Count > 0) node.Content = content;
        return node;
    }

    /// <summary>Caixas de texto sem âncora: o conteúdo entra na linha.</summary>
    private IEnumerable<Node> ReadTextBoxesInline(OpenXmlElement shape)
    {
        foreach (var box in OutermostTextBoxes(shape))
        {
            // Aqui o texto entra na linha e a caixa não é desenhada em lugar
            // nenhum: qualquer moldura que ela tenha se perde de vista, e é
            // disso que o aviso fala.
            var look = ShapeLook.Of(box);
            if (!look.Complete || look.Draws) inventory.NoteInvisible(Inventory.Shapes);

            foreach (var paragraph in box.Descendants<Paragraph>())
            {
                if (paragraph.Ancestors<TextBoxContent>().First() != box) continue;

                var (_, inheritedRun) = _styles.Resolve(paragraph.ParagraphProperties);
                var afterSomething = _paragraphHasContent;

                var outer = _directRuns;
                _directRuns = false;
                var inline = ReadInline(paragraph, inheritedRun);
                _directRuns = outer;
                if (inline.Count == 0) continue;

                if (afterSomething) yield return Node.Of("hardBreak");
                _paragraphHasContent = true;

                foreach (var node in inline) yield return node;
            }
        }
    }

    /// <summary>
    /// As caixas de texto mais externas de <paramref name="root"/>.
    /// </summary>
    /// <remarks>
    /// Para na primeira caixa de cada ramo em vez de usar
    /// <c>Descendants</c>: uma caixa dentro de outra é alcançada pela recursão
    /// de <see cref="ReadInline"/>, e as duas rotas juntas escreveriam o texto
    /// de dentro duas vezes.
    /// </remarks>
    private static IEnumerable<TextBoxContent> OutermostTextBoxes(OpenXmlElement root) =>
        TextBoxNav.Outermost(root);

    // --- imagens ------------------------------------------------------------

    /// <summary>
    /// Imagem → nó com data URI.
    /// </summary>
    /// <remarks>
    /// Vale para a imagem no meio da linha (`wp:inline`) e também para a
    /// ancorada que está onde o fluxo já a poria — o jeito do LibreOffice
    /// gravar "imagem no próprio parágrafo". Quem separa as duas é
    /// <see cref="AnchorReader.FlowsWithText"/>. Ver docs/01-corpus-docx.md,
    /// Descoberta 3.
    /// </remarks>
    /// <summary>Os bytes da imagem como data URI, ou `null` se não houver.</summary>
    /// <remarks>
    /// Separado de <see cref="ReadImage"/> porque o objeto ancorado precisa dos
    /// bytes sem o nó: ele não vira bloco no fluxo, vira posição na folha.
    /// </remarks>
    private string? ImageSourceOf(OpenXmlElement drawing)
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

        return $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}";
    }

    /// <summary>
    /// A captura ancorada ao parágrafo vai para o começo dele.
    /// </summary>
    /// <remarks>
    /// A âncora diz "no alto do parágrafo" (deslocamento vertical zero a partir
    /// dele), e o run do desenho pode estar em qualquer ponto da frase: no corpus,
    /// "Múltiplos" + captura + " Registros do InfPercurso OK". O Word e o
    /// LibreOffice põem o quadro no topo e o texto embaixo; desenhá-lo na ordem
    /// dos runs deixava uma linha de texto acima do quadro — 22 pt a mais na
    /// folha, e as folhas seguintes cortando uma captura antes.
    ///
    /// A ordem dentro do parágrafo não muda o desenho do Word para a âncora
    /// relativa ao parágrafo, e o parágrafo que ninguém edita volta do original,
    /// byte a byte. Ancorada à **linha** fica onde está: aí o lugar do run é a
    /// posição.
    /// </remarks>
    private void TopAnchoredFirst(List<Node> content)
    {
        if (_topAnchored.Count == 0) return;
        var moved = content.Where(_topAnchored.Contains).ToList();
        if (moved.Count == 0 || content.Take(moved.Count).SequenceEqual(moved)) return;
        content.RemoveAll(_topAnchored.Contains);
        content.InsertRange(0, moved);
    }

    private Node? ReadImage(OpenXmlElement drawing)
    {
        if (ImageSourceOf(drawing) is not { } src) return null;

        var node = Node.Of("image").With("src", src);

        // Ancorada ao parágrafo, ainda que no lugar em que o fluxo já a poria.
        //
        // A diferença é de altura, e ela conta: o parágrafo que **ancora** uma
        // imagem ocupa a linha dele além da imagem, e o que a traz no meio da
        // frase ocupa só a imagem. Medido no LibreOffice: 11,55 pt entre duas
        // capturas seguidas de um documento de evidências, que é exatamente uma
        // linha de Arial 10 pt.
        if (AnchorReader.AnchorOf(drawing) is { } anchor)
        {
            node.With("anchored", true);
            var vertical = anchor.GetFirstChild<Drawing.Wordprocessing.VerticalPosition>();
            var from = vertical?.RelativeFrom?.Value;
            // Só no topo do parágrafo: deslocamento zero — até 1 pt, que é o
            // arredondamento com que o LibreOffice grava o zero (635 EMU no
            // corpus). Deslocado para baixo, o quadro pode ter texto acima dele,
            // e o run fica onde está.
            var offset = long.TryParse(vertical?.PositionOffset?.Text, out var emus) ? Math.Abs(emus) : 0;
            if ((from is null || from == Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph) &&
                offset <= 12700)
            {
                _topAnchored.Add(node);
            }
        }

        // O texto alternativo, que é o que um leitor de tela lê no lugar da
        // imagem. Mora no `wp:docPr/@descr` e chega ao editor como o `alt` do
        // `<img>` — os dois existem para a mesma pessoa. Só quando há algo
        // escrito: `descr=""` é o padrão de quem nunca preencheu o campo, e
        // emiti-lo como atributo faria toda imagem divergir do que o editor
        // devolve.
        var properties = drawing.Descendants<Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
        if (properties?.Description?.Value is { Length: > 0 } description)
        {
            node.With("alt", description);
        }

        var extent = drawing.Descendants<Drawing.Wordprocessing.Extent>().FirstOrDefault();
        if (extent?.Cx?.Value is { } wide && extent.Cy?.Value is { } tall && wide > 0 && tall > 0)
        {
            // `wp:extent` mede a imagem antes de girar. Num quarto de volta o
            // que ocupa a largura da página é a altura dela, e usar `cx` põe na
            // linha uma imagem deitada com a medida do lado comprido: a marca
            // vertical de 28,58 cm da capa do modelo de manual chegava como
            // 1080 px de largura numa coluna de 734 px, tomava a página inteira
            // e empurrava o resto para baixo.
            var (across, down) = IsQuarterTurned(drawing) ? (tall, wide) : (wide, tall);

            // As duas medidas, e não só a largura. A altura faltando tinha três
            // consequências, todas silenciosas: o navegador reservava zero até a
            // imagem decodificar, e a paginação media a folha sem ela; a
            // proporção passava a ser a do arquivo, e não a que o documento
            // pede, então imagem esticada de propósito voltava ao natural; e na
            // gravação o escritor chutava três quartos da largura, o que dava
            // outro tamanho ao que ninguém tinha tocado.
            //
            // EMU → pixels CSS: 914400 EMU por polegada, 96 px por polegada.
            node.With("width", (int)Math.Round(across * 96.0 / 914400));
            node.With("height", (int)Math.Round(down * 96.0 / 914400));
        }

        return node;
    }

    /// <summary>
    /// A imagem está girada perto de um quarto de volta, para um lado ou para
    /// o outro?
    /// </summary>
    /// <remarks>
    /// `a:rot` vem em 60000 avos de grau e pode ser negativo. Só o quarto de
    /// volta interessa aqui, porque é o único ângulo em que largura e altura
    /// trocam de papel; um giro pequeno mantém a medida aproximadamente igual e
    /// não vale a conta do retângulo envolvente.
    /// </remarks>
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

        // A largura de cada coluna vem da grade da tabela, uma vez só: é ela que
        // o Word usa para desenhar e é ela que o editor espelha no `colgroup`.
        var grid = TableLook.GridWidths(table);

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = new List<Node>();

            // `w:tblHeader` é a linha que o Word repete no alto de cada página, e
            // é exatamente o que o editor chama de linha de cabeçalho. Sem esta
            // leitura ela chegava como linha comum, e ligar a linha de cabeçalho
            // na tela de um documento que já a tinha a **desligava** no arquivo.
            // Presente sem `w:val` já é "sim"; só `w:val="false"` nega.
            var repeat = row.TableRowProperties?.GetFirstChild<TableHeader>();
            var header = repeat is not null
                         && !(repeat.Val is { } declared && declared.Value == OnOffOnlyValues.Off);

            // A linha pode começar colunas adiante — `w:gridBefore`, comum em
            // formulário e em documento convertido de PDF. Sem pular essas
            // colunas, cada célula ganhava a largura da coluna à esquerda da dela.
            var column = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value is > 0 and var before
                ? before
                : 0;

            foreach (var cell in row.Elements<TableCell>())
            {
                // Tabela dentro de tabela é comum em documento de formulário, e
                // ler só os parágrafos a fazia desaparecer da tela — com o texto
                // dentro dela. O `tableCell` do editor aceita bloco, e a tabela
                // aninhada é um bloco.
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

                // Os dois são **sempre** escritos, pelo mesmo motivo do `indent`
                // do parágrafo: o editor declara `colspan` e `rowspan` com padrão
                // 1 e devolve os dois em toda célula. Omitidos aqui, os dois lados
                // descreviam a mesma célula de formas diferentes e a comparação
                // que decide o que preservar dizia "mudou" em tabela que ninguém
                // tocou — toda tabela era regenerada ao salvar, e nada falhava.
                var span = cell.TableCellProperties?.GridSpan?.Val?.Value;
                var spanned = span is > 1 ? span.Value : 1;
                node.With("colspan", spanned);

                // O modelo do editor não representa mesclagem vertical: ela fica
                // no `w:vMerge` do arquivo, que a gravação devolve ao lugar.
                node.With("rowspan", 1);

                // A largura das colunas que esta célula ocupa, na forma que o
                // TableKit usa: uma medida por coluna da grade, em pixels do CSS.
                // Enquanto ela não era lida, arrastar a divisória escrevia o
                // número novo no modelo e o gravador devolvia a grade antiga ao
                // arquivo — perda silenciosa.
                if (grid.Count >= column + spanned)
                {
                    var widths = new JsonArray();
                    for (var index = 0; index < spanned; index++) widths.Add(grid[column + index]);
                    node.With("colwidth", widths);
                }

                column += spanned;

                // Sombreamento e bordas da célula. Só quando o arquivo os declara:
                // atributo ausente e atributo nulo são a mesma afirmação, e é a
                // que mantém a impressão digital igual à do editor.
                if (TableLook.Shading(cell.TableCellProperties) is { } fill) node.With("shading", fill);
                if (TableLook.Borders(cell.TableCellProperties) is { } borders) node.With("borders", borders);

                cells.Add(node);
            }

            var rowNode = Node.Of("tableRow");
            rowNode.Content = cells;
            rows.Add(rowNode);
        }

        var tableNode = Node.Of("table");
        tableNode.Content = rows;

        // A margem de célula que o Word usa nesta tabela, para a tela medir a
        // linha como o papel. Só leitura: o `w:tblPr` original volta como estava.
        tableNode.With("cellMargins", string.Join(' ', TableLook.CellMargins(table, part)));
        return tableNode;
    }

}

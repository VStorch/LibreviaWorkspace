using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Nó do editor → OOXML. Só roda para blocos que o usuário **editou**.
/// </summary>
/// <remarks>
/// O que este arquivo não souber gerar é perda de verdade, e por isso ele é o
/// lugar mais perigoso da Fase 4. A edição cirúrgica reduz o estrago: blocos
/// intactos nunca passam por aqui — vão direto do arquivo original para o novo.
/// </remarks>
public sealed class ParagraphWriter
{
    private readonly MainDocumentPart _part;
    private readonly Inventory _inventory;
    private readonly ParagraphFormat _format;
    private readonly StyleResolver _styles;

    /// <summary>
    /// O que o estilo do parágrafo em gravação dá aos runs — ver
    /// <see cref="DropWhatRepeatsTheStyle"/>.
    /// </summary>
    private RunProperties _paragraphRun = new();
    private readonly TableWriter _tables;
    private readonly ImageWriter _images;

    /// <param name="usableWidthPx">
    /// A largura da coluna de texto, em pixels do CSS. É o teto de uma imagem
    /// que chega sem medida: maior do que isso, o Word a desenha estourando a
    /// margem.
    /// </param>
    internal ParagraphWriter(
        MainDocumentPart part,
        Inventory inventory,
        int usableWidthPx = ImageWriter.DefaultWidthPx,
        HeadingStyles? headings = null,
        bool flatten = false)
    {
        _part = part;
        _inventory = inventory;
        _styles = new StyleResolver(part);
        _format = new ParagraphFormat(inventory, headings ?? new HeadingStyles(part, null), _styles, flatten);
        var usable = usableWidthPx > 0 ? usableWidthPx : ImageWriter.DefaultWidthPx;
        _tables = new TableWriter(inventory, (node, original) => Write(node, null, original), usable);
        _images = new ImageWriter(part, inventory, usable);
    }

    /// <param name="original">
    /// O parágrafo como estava no arquivo, quando existe.
    /// </param>
    public IEnumerable<OpenXmlElement> Write(
        Node node,
        ListPlacement? list = null,
        OpenXmlElement? original = null)
    {
        switch (node.Type)
        {
            case "paragraph":
            case "heading":
            {
                var paragraph = WriteParagraph(node, list, original as Paragraph);
                CarryAnchored(paragraph, node, original);
                yield return paragraph;
                break;
            }

            case "pageBreak":
                yield return new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                break;

            case "table":
                yield return _tables.Write(node, original as Table);
                break;

            // A imagem que a pessoa insere pela barra de ferramentas é um bloco,
            // e não um trecho de linha: no OOXML não existe imagem fora de
            // parágrafo, então ela viaja dentro de um. Sem este caso ela caía no
            // ramo de baixo — parágrafo vazio e um aviso de perda — e desaparecia
            // do documento ao salvar.
            case "image":
            {
                var image = _images.Write(node);
                if (image is null)
                {
                    yield return new Paragraph();
                    break;
                }

                // O alinhamento da imagem é o `w:jc` do parágrafo que a carrega —
                // no OOXML não existe "imagem centralizada", existe parágrafo
                // centralizado com uma imagem dentro. A imagem lida de um `.docx`
                // já vem dentro de um parágrafo, e ali quem alinha é o próprio
                // parágrafo; este caso é o da imagem inserida pela barra, que é um
                // bloco do editor e não tem parágrafo seu.
                var aligned = new Paragraph(image);
                if (ParagraphFormat.JustificationOf(Attr.String(node, "align")) is { } justification)
                {
                    aligned.ParagraphProperties = new ParagraphProperties(justification);
                }

                yield return aligned;
                break;
            }

            case "horizontalRule":
                yield return new Paragraph(new ParagraphProperties(
                    new ParagraphBorders(new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 6,
                        Color = "auto",
                    })));
                break;

            default:
                // Um bloco que não sabemos gerar não pode virar nada em
                // silêncio: vira parágrafo vazio e entra no inventário.
                _inventory.NoteLoss($"bloco do tipo \"{node.Type}\"");
                yield return new Paragraph();
                break;
        }
    }

    /// <summary>Numeração herdada do documento, para itens de lista.</summary>
    /// <param name="Kind">
    /// `bulletList` ou `orderedList`. Viaja junto porque uma sublista só pode
    /// herdar a numeração da lista de fora quando é do mesmo tipo: herdada às
    /// cegas, uma sublista numerada dentro de uma com marcador sai com marcador.
    /// </param>
    public sealed record ListContext(string Kind, int NumberingId, int Level);

    /// <summary>
    /// O que quem chama sabe sobre a numeração do parágrafo.
    /// </summary>
    /// <remarks>
    /// São três estados, e não dois. O corpo sabe que o parágrafo **é** item de
    /// lista — o embrulho com o contexto dentro; sabe que **não é** — o embrulho
    /// com <c>null</c> dentro, e aí o `w:numPr` do original tem de sair, senão o
    /// parágrafo continua numerado depois de a pessoa ter tirado a lista; e a
    /// tabela **não sabe** — nem embrulho, e aí o `w:numPr` do original fica.
    ///
    /// O terceiro caso não existia, e a célula pagava por isso: a detecção de
    /// lista mora no laço do corpo, nunca dentro da célula, de modo que a tabela
    /// chamava o escritor sempre sem contexto. Corrigir uma palavra numa tabela
    /// tirava os marcadores da lista que estava na célula, sem nada no inventário.
    /// </remarks>
    public readonly record struct ListPlacement(ListContext? List);

    /// <summary>
    /// Os objetos ancorados do parágrafo original seguem no parágrafo reescrito.
    /// </summary>
    /// <remarks>
    /// A capa do modelo de manual é uma forma só: título e subtítulo moram em
    /// caixas de texto ancoradas, e a marca lateral é uma imagem posicionada.
    /// Nada disso este escritor sabe gerar do zero — o que ele tem é o XML
    /// original, e copiá-lo é o mesmo remédio que a gravação cirúrgica usa para
    /// o bloco inteiro, aplicado a um pedaço dele.
    ///
    /// Sem isto, editar o parágrafo que ancora uma forma a apagava, e era esse
    /// risco que mantinha o documento inteiro em somente leitura.
    ///
    /// Só o `w:r` que **é** o desenho: um run que também traga texto seria
    /// copiado com o texto junto, e a frase apareceria duas vezes. Esse caso
    /// continua entrando no inventário como perda.
    /// </remarks>
    private void CarryAnchored(Paragraph paragraph, Node node, OpenXmlElement? original)
    {
        if (original is null) return;

        // O ancorado que corre com o texto chegou ao editor como imagem do
        // parágrafo, e é por ela que ele volta. Copiá-lo aqui também punha duas
        // imagens no arquivo — e trazia de volta a que a pessoa tivesse apagado.
        var flowing = ImageWriter.FlowingImagesOf(original);

        foreach (var run in original.Elements<Run>())
        {
            if (!IsAnchoredOnly(run)) continue;
            if (run.Elements<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any(flowing.Contains)) continue;
            paragraph.AppendChild(run.CloneNode(true));
        }

        ApplyBoxText(paragraph, node);
    }

    /// <summary>
    /// O que a pessoa digitou dentro de uma caixa volta para o `w:txbxContent`.
    /// </summary>
    /// <remarks>
    /// A capa do modelo de manual é feita disto: título e subtítulo moram em
    /// caixas, e sem esta volta a caixa seria editável na tela e voltaria com o
    /// texto antigo no arquivo — pior do que não deixar editar.
    ///
    /// Caixa cujo texto não mudou **não é tocada**: o XML dela segue como
    /// estava, com a moldura, o preenchimento e a formatação que este escritor
    /// não sabe reproduzir. É a mesma aposta da gravação cirúrgica, um nível
    /// abaixo.
    ///
    /// Se a contagem não bater — um desenho que não pôde ser copiado, uma caixa
    /// que o leitor viu e o escritor não — nada é escrito: acertar a caixa
    /// errada poria o subtítulo dentro do título, e um texto perdido é menos
    /// grave do que um texto trocado de lugar.
    /// </remarks>
    private void ApplyBoxText(Paragraph paragraph, Node node)
    {
        var wanted = BoxContentsOf(node);
        if (wanted.Count == 0) return;

        var boxes = TextBoxNav.AnchoredBoxesOf(paragraph).ToList();
        if (boxes.Count != wanted.Count)
        {
            _inventory.NoteLoss("texto de caixa num parágrafo que você editou");
            return;
        }

        var touched = false;
        for (var index = 0; index < boxes.Count; index++)
        {
            var box = boxes[index];
            var content = wanted[index];
            if (TextBoxNav.TextOf(box) == PlainTextOf(content)) continue;

            box.RemoveAllChildren();
            foreach (var block in content)
            {
                foreach (var element in Write(block)) box.AppendChild(element);
            }

            // `w:txbxContent` vazio invalida o documento para o Word.
            if (!box.HasChildren) box.AppendChild(new Paragraph());
            touched = true;
        }

        if (!touched) return;

        foreach (var alternate in paragraph.Descendants<AlternateContent>().ToList())
        {
            TextBoxNav.MirrorFallback(alternate);
        }
    }

    /// <summary>O conteúdo das caixas de texto do bloco, na ordem do modelo.</summary>
    private static List<List<Node>> BoxContentsOf(Node node)
    {
        var contents = new List<List<Node>>();
        if (node.Attrs is null ||
            !node.Attrs.TryGetValue("floats", out var value) ||
            value is not JsonArray floats)
        {
            return contents;
        }

        foreach (var item in floats)
        {
            if (item is not JsonObject float_) continue;
            if (float_["kind"]?.GetValue<string>() != "text") continue;

            var content = float_["content"].Deserialize<List<Node>>(DocxJson.Options);
            contents.Add(content ?? []);
        }

        return contents;
    }

    /// <summary>O texto de um conteúdo de caixa, na mesma forma que o do arquivo.</summary>
    private static string PlainTextOf(List<Node> content) =>
        string.Join("\n", content.Select(TextOfNode));

    private static string TextOfNode(Node node) =>
        node.Text ?? string.Concat((node.Content ?? []).Select(TextOfNode));

    /// <summary>O run carrega um objeto ancorado e mais nada que se escreva.</summary>
    /// <remarks>
    /// Pelos filhos diretos, e não por `Descendants`: o texto de dentro de uma
    /// caixa também é `w:t`, e procurá-lo em profundidade rejeitava justamente
    /// os runs que existem para ser copiados.
    /// </remarks>
    internal static bool IsAnchoredOnly(Run run) =>
        run.Descendants<WordDrawing.Anchor>().Any() && !run.Elements<Text>().Any();

    private Paragraph WriteParagraph(Node node, ListPlacement? list, Paragraph? original)
    {
        var paragraph = new Paragraph();

        // O `w:pPr` sai de ParagraphFormat, que parte do original: o estilo, o
        // espaçamento, a entrelinha, o fundo, a marca de parágrafo e o `w:sectPr`
        // do arquivo sobrevivem à edição porque ninguém os reescreve.
        if (_format.Build(node, list, original) is { } properties)
        {
            paragraph.ParagraphProperties = properties;
        }

        // Os marcadores que abriam o parágrafo abrem o parágrafo reescrito.
        foreach (var mark in Bookmarks(original, leading: true))
        {
            paragraph.AppendChild(mark.CloneNode(true));
        }

        _paragraphRun = _styles.Resolve(paragraph.ParagraphProperties).Run;

        // As imagens que já estavam no parágrafo voltam com o desenho original —
        // ver ImageWriter.Reuse.
        var images = ImageWriter.FlowingImagesOf(original);
        foreach (var child in node.Content ?? [])
        {
            foreach (var element in WriteInline(child, images)) paragraph.AppendChild(element);
        }

        // A quebra volta para onde estava: no fim do parágrafo, dentro de um
        // `w:r`. O leitor a transformou em propriedade do bloco para não pôr um
        // nó de bloco em posição de linha; aqui ela desfaz o caminho. Sem isto,
        // editar o parágrafo que carrega a quebra a apagaria em silêncio.
        if (Attr.Bool(node, "breakAfter"))
        {
            paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        }

        foreach (var mark in Bookmarks(original, leading: false))
        {
            paragraph.AppendChild(mark.CloneNode(true));
        }

        return paragraph;
    }

    /// <summary>
    /// Os marcadores do parágrafo original, separados pelo lado em que estavam.
    /// </summary>
    /// <remarks>
    /// `w:bookmarkStart` e `w:bookmarkEnd` são o destino da referência cruzada, da
    /// entrada de índice e do link interno do documento — e o modelo do editor não
    /// os representa. Reescrevendo o parágrafo só a partir do modelo, eles
    /// desapareciam do arquivo: quem os citava passava a apontar para o vazio, e
    /// nada disso chegava ao inventário. É a mesma solução dos objetos ancorados:
    /// o que o editor não sabe dizer vem do XML original.
    ///
    /// Pelo lado em que estavam, e não todos juntos num canto: um marcador que
    /// abraça o parágrafo tem o começo antes do texto e o fim depois dele, e
    /// levar os dois para o mesmo lado encurtaria o trecho marcado até o vazio.
    /// A posição **dentro** do texto não sobrevive — o modelo não diz onde o
    /// marcador começava no meio da frase —, e é a perda que resta: o marcador
    /// continua existindo e continua neste parágrafo.
    ///
    /// Marcador é conteúdo de nível de run, então ele cabe em qualquer ponto do
    /// `w:p` depois do `w:pPr`.
    /// </remarks>
    /// <param name="leading">
    /// Verdadeiro para os que vinham antes de qualquer conteúdo, falso para os
    /// demais.
    /// </param>
    private static IEnumerable<OpenXmlElement> Bookmarks(Paragraph? original, bool leading)
    {
        if (original is null) yield break;

        var started = false;
        foreach (var child in original.ChildElements)
        {
            if (child is BookmarkStart or BookmarkEnd)
            {
                if (started != leading) yield return child;
                continue;
            }

            if (child is not ParagraphProperties) started = true;
        }
    }

    private IEnumerable<OpenXmlElement> WriteInline(
        Node node,
        List<DocumentFormat.OpenXml.Wordprocessing.Drawing>? originalImages = null)
    {
        switch (node.Type)
        {
            case "text":
                yield return WriteTextRun(node);
                break;

            case "hardBreak":
                yield return new Run(new Break());
                break;

            case "pageBreak":
                yield return new Run(new Break { Type = BreakValues.Page });
                break;

            case "image":
                if ((originalImages is null ? null : _images.Reuse(node, originalImages)) is { } kept)
                {
                    yield return kept;
                }
                else if (_images.Write(node) is { } image)
                {
                    yield return image;
                }

                break;

            default:
                _inventory.NoteLoss($"conteúdo do tipo \"{node.Type}\"");
                break;
        }
    }

    private OpenXmlElement WriteTextRun(Node node)
    {
        var run = new Run();
        var properties = new RunProperties();
        string? hyperlink = null;

        foreach (var mark in node.Marks ?? [])
        {
            switch (mark.Type)
            {
                // `off` é o trecho desligando o que o estilo do parágrafo liga —
                // ver RunReader.Off.
                case "bold": properties.Bold = IsOff(mark) ? new Bold { Val = false } : new Bold(); break;
                case "italic": properties.Italic = IsOff(mark) ? new Italic { Val = false } : new Italic(); break;
                case "strike": properties.Strike = IsOff(mark) ? new Strike { Val = false } : new Strike(); break;
                case "caps": properties.Caps = new Caps(); break;
                case "smallCaps": properties.SmallCaps = new SmallCaps(); break;
                case "underline":
                    properties.Underline = new Underline
                    {
                        Val = IsOff(mark) ? UnderlineValues.None : UnderlineValues.Single,
                    };
                    break;

                // Uma propriedade só para os dois, com valores que se excluem —
                // é assim no OOXML, e é assim no editor, onde as duas marcas se
                // excluem uma à outra. Enquanto não estavam aqui, o expoente
                // sobrevivia à leitura e morria na gravação: o trecho voltava
                // para a linha do texto e o aviso saía no inventário.
                case "superscript":
                    properties.VerticalTextAlignment = new VerticalTextAlignment
                    {
                        Val = VerticalPositionValues.Superscript,
                    };
                    break;

                case "subscript":
                    properties.VerticalTextAlignment = new VerticalTextAlignment
                    {
                        Val = VerticalPositionValues.Subscript,
                    };
                    break;

                case "highlight":
                    if (Attr.MarkString(mark, "color") is { } fill) ApplyHighlight(properties, fill);
                    break;

                case "link":
                    hyperlink = Attr.MarkString(mark, "href");
                    break;

                case "charStyle":
                    if (Attr.MarkString(mark, "styleId") is { Length: > 0 } characterStyle)
                    {
                        properties.RunStyle = new RunStyle { Val = characterStyle };
                    }

                    break;

                case "textStyle":
                    ApplyTextStyle(properties, mark);
                    break;

                default:
                    _inventory.NoteLoss($"formatação \"{mark.Type}\"");
                    break;
            }
        }

        DropWhatRepeatsTheStyle(properties, _paragraphRun);
        if (properties.HasChildren) run.RunProperties = properties;

        // O texto entra peça por peça: tabulação é `w:tab`, quebra de linha é
        // `w:br`, e caractere de controle não existe no XML 1.0 — escrevê-lo
        // derrubava a gravação inteira. Ver XmlText.
        var pieces = XmlText.Of(node.Text).ToList();
        if (pieces.Count == 0)
        {
            // Run sem nada dentro é inválido para o Word.
            pieces.Add(new Text(string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        }

        foreach (var piece in pieces) run.AppendChild(piece);

        if (hyperlink is null) return run;

        // `w:hyperlink` embrulha o run — não cabe dentro dele.
        Uri target;
        try
        {
            target = new Uri(hyperlink, UriKind.Absolute);
        }
        catch (UriFormatException)
        {
            _inventory.NoteLoss("endereço de link inválido");
            return run;
        }

        var relationship = _part.AddHyperlinkRelationship(target, true);
        var link = new Hyperlink { Id = relationship.Id };
        link.AppendChild(run);
        return link;
    }

    /// <summary>
    /// Tira do run o que o estilo do parágrafo já lhe dá.
    /// </summary>
    /// <remarks>
    /// As marcas do editor chegam achatadas — o leitor põe em cada trecho a fonte,
    /// o tamanho e o negrito do estilo —, e gravá-las de volta como formatação
    /// direta repetia o estilo em cada `w:r` do bloco editado. O arquivo abria
    /// igual, mas desligado do estilo: mudar a fonte do Normal no Word não
    /// alcançava mais o parágrafo que alguém tinha corrigido aqui.
    ///
    /// Só o que coincide sai. Marca ausente não vira "desligado" explícito: na tela
    /// o trecho sem marca mostra o estilo, e é o estilo que ele continua a ter.
    /// </remarks>
    private static void DropWhatRepeatsTheStyle(RunProperties properties, RunProperties style)
    {
        // Ligado ou desligado, o que coincide com o estilo sai: o "desligado"
        // num estilo que não liga nada não diz nada.
        if (properties.Bold is not null && RunReader.IsOn(properties.Bold) == RunReader.IsOn(style.Bold))
        {
            properties.Bold = null;
        }

        if (properties.Italic is not null && RunReader.IsOn(properties.Italic) == RunReader.IsOn(style.Italic))
        {
            properties.Italic = null;
        }

        if (properties.Strike is not null && RunReader.IsOn(properties.Strike) == RunReader.IsOn(style.Strike))
        {
            properties.Strike = null;
        }

        if (properties.Caps is not null && RunReader.IsOn(style.Caps)) properties.Caps = null;
        if (properties.SmallCaps is not null && RunReader.IsOn(style.SmallCaps)) properties.SmallCaps = null;

        if (properties.Underline?.Val is { } underline &&
            (underline.Value != UnderlineValues.None) ==
            (style.Underline?.Val is { } line && line.Value != UnderlineValues.None))
        {
            properties.Underline = null;
        }

        if (Same(properties.RunFonts?.Ascii?.Value, style.RunFonts?.Ascii?.Value)) properties.RunFonts = null;
        if (Same(properties.FontSize?.Val?.Value, style.FontSize?.Val?.Value)) properties.FontSize = null;
        if (Same(properties.Color?.Val?.Value, style.Color?.Val?.Value)) properties.Color = null;
        if (Same(properties.VerticalTextAlignment?.Val?.InnerText, style.VerticalTextAlignment?.Val?.InnerText))
        {
            properties.VerticalTextAlignment = null;
        }
    }

    private static bool IsOff(Mark mark) =>
        mark.Attrs?.TryGetValue("off", out var off) == true &&
        off?.GetValueKind() == System.Text.Json.JsonValueKind.True;

    private static bool Same(string? a, string? b) =>
        a is not null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private void ApplyTextStyle(RunProperties properties, Mark mark)
    {
        if (Attr.MarkString(mark, "color") is { } color)
        {
            if (ColorValue.Hex(color) is { } hex) properties.Color = new Color { Val = hex };
            else _inventory.NoteLoss($"cor de texto \"{color}\"");
        }

        // O fundo de um trecho de texto: no editor é `backgroundColor`, no
        // arquivo é o mesmo `w:shd` do realce. Enquanto ficava de fora, pintar o
        // fundo de uma palavra não chegava ao documento.
        if (Attr.MarkString(mark, "backgroundColor") is { } background)
        {
            ApplyHighlight(properties, background);
        }

        if (Attr.MarkString(mark, "fontFamily") is { } font &&
            ParagraphFormat.FirstFont(font) is { } first)
        {
            properties.RunFonts = new RunFonts { Ascii = first, HighAnsi = first };
        }

        if (Attr.MarkString(mark, "fontSize") is { } size)
        {
            // `w:sz` é em meios-pontos, e a medida pode chegar em pixels: o
            // editor grava `font-size` como o CSS o escreve, e um `16px` lido
            // como "16 pt" engordava o texto em um terço.
            if (Attr.Points(size) is { } points && points > 0)
            {
                var halfPoints = (int)Math.Round(points * 2);
                properties.FontSize = new FontSize
                {
                    Val = halfPoints.ToString(CultureInfo.InvariantCulture),
                };
            }
            else
            {
                _inventory.NoteLoss($"tamanho de fonte \"{size}\"");
            }
        }

        // A entrelinha é propriedade do parágrafo no OOXML: não existe `w:line`
        // dentro de um `w:rPr`. Quem a grava é ParagraphFormat, a partir do
        // atributo do bloco; aplicada a um trecho só, ela não tem para onde ir.
        if (Attr.MarkString(mark, "lineHeight") is { } lineHeight)
        {
            _inventory.NoteLoss($"entrelinha de um trecho de texto (\"{lineHeight}\")");
        }
    }

    /// <summary>
    /// Fundo de texto, como `w:shd`.
    /// </summary>
    /// <remarks>
    /// Hexadecimal de seis dígitos, sempre: o editor guarda a cor como o CSS a
    /// escreve — `rgb(255, 0, 0)`, `#f00`, `red` — e o atributo do OOXML não
    /// aceita nenhuma dessas formas. `rgb(...)` no lugar fazia o Word declarar o
    /// documento danificado; um nome de cor era aceito e desenhado como preto.
    /// </remarks>
    private void ApplyHighlight(RunProperties properties, string color)
    {
        if (ColorValue.Hex(color) is not { } hex)
        {
            _inventory.NoteLoss($"cor de fundo de texto \"{color}\"");
            return;
        }

        properties.Shading = new Shading
        {
            Val = ShadingPatternValues.Clear,
            Color = "auto",
            Fill = hex,
        };
    }
}

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
/// Nó do editor → OOXML, só para blocos **editados**: o que este arquivo não sabe
/// gerar é perda de verdade.
/// </summary>
public sealed class ParagraphWriter
{
    private readonly MainDocumentPart _part;

    /// <summary>O documento, ou <c>footnotes.xml</c> na nota: o <c>r:id</c> do link de uma nota é dela.</summary>
    private readonly OpenXmlPart _owner;

    /// <summary>Ver NotesWriter.ReferenceRunsOf.</summary>
    private Dictionary<string, Run>? _noteRuns;

    private readonly Inventory _inventory;
    private readonly ParagraphFormat _format;
    private readonly StyleResolver _styles;

    /// <summary>Ver <see cref="DropWhatRepeatsTheStyle"/>.</summary>
    private RunProperties _paragraphRun = new();
    private readonly TableWriter _tables;
    private readonly ImageWriter _images;

    /// <summary>Desligado, no rascunho anterior aos marcadores, os do parágrafo original são copiados.</summary>
    private readonly bool _references;

    /// <summary>Lido na primeira ponta, antes de o corpo ser trocado — ver <see cref="CommentAnchors"/>.</summary>
    private CommentAnchors? _comments;

    /// <summary>Em vinte-avos de ponto: 1 px do CSS são 15.</summary>
    private readonly int _usableTwips;

    /// <param name="usableWidthPx">O teto, em pixels do CSS, da imagem que chega sem medida.</param>
    internal ParagraphWriter(
        MainDocumentPart part,
        Inventory inventory,
        int usableWidthPx = ImageWriter.DefaultWidthPx,
        HeadingStyles? headings = null,
        bool flatten = false,
        bool references = true,
        bool revisions = true,
        OpenXmlPart? owner = null)
    {
        _part = part;
        _owner = owner ?? part;
        _references = references;
        _inventory = inventory;
        _styles = new StyleResolver(part);
        _format = new ParagraphFormat(inventory, headings ?? new HeadingStyles(part, null), _styles, flatten, revisions);
        var usable = usableWidthPx > 0 ? usableWidthPx : ImageWriter.DefaultWidthPx;
        _tables = new TableWriter(inventory, (node, original) => Write(node, null, original), usable, revisions);
        _images = new ImageWriter(part, inventory, usable, _owner);
        _usableTwips = usable * 15;
    }

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
                NoteEditedMathFormatting(node, original);
                yield return paragraph;
                break;
            }

            case "pageBreak":
                yield return new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                break;

            case "table":
                yield return _tables.Write(node, original as Table);
                break;

            // A imagem inserida pela barra é um bloco do editor, e no OOXML viaja dentro de um parágrafo.
            case "image":
            {
                var image = _images.Write(node);
                if (image is null)
                {
                    yield return new Paragraph();
                    break;
                }

                // Quem alinha é o `w:jc` do parágrafo: no OOXML não há imagem centralizada.
                var aligned = new Paragraph(image);
                if (ParagraphFormat.JustificationOf(Attr.String(node, "align")) is { } justification)
                {
                    aligned.ParagraphProperties = new ParagraphProperties(justification);
                }

                yield return aligned;
                break;
            }

            case "tableOfContents":
                foreach (var element in WriteTableOfContents(node, original)) yield return element;
                break;

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
                // O que não sabemos gerar vira parágrafo vazio e entra no inventário, nunca em silêncio.
                _inventory.NoteLoss($"bloco do tipo \"{node.Type}\"");
                yield return new Paragraph();
                break;
        }
    }

    /// <summary>Numeração herdada do documento, para itens de lista.</summary>
    /// <param name="Kind">A sublista só herda a numeração da de fora quando é do mesmo tipo.</param>
    public sealed record ListContext(string Kind, int NumberingId, int Level);

    /// <summary>
    /// Três estados: o corpo sabe que o parágrafo **é** item (o contexto), que **não
    /// é** (<c>null</c>, e o <c>w:numPr</c> do original sai), e a tabela **não sabe** (sem
    /// embrulho, e o <c>w:numPr</c> fica).
    /// </summary>
    public readonly record struct ListPlacement(ListContext? List);

    /// <summary>
    /// Os objetos ancorados do original seguem no parágrafo reescrito, como a
    /// gravação cirúrgica faz com o bloco inteiro: este escritor não os gera. Só o
    /// <c>w:r</c> que **é** o desenho; um run com texto repetiria a frase e vai ao
    /// inventário.
    /// </summary>
    private void CarryAnchored(Paragraph paragraph, Node node, OpenXmlElement? original)
    {
        if (original is null) return;

        // O ancorado que corre com o texto volta como imagem do parágrafo; copiado aqui, viria duas vezes.
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
    /// O texto digitado numa caixa volta ao <c>w:txbxContent</c>. Caixa sem mudança
    /// não é tocada. Se a contagem não bater, nada é escrito: um texto trocado de
    /// caixa é pior que um perdido.
    /// </summary>
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

            // `w:txbxContent` vazio invalida o documento.
            if (!box.HasChildren) box.AppendChild(new Paragraph());
            touched = true;
        }

        if (!touched) return;

        foreach (var alternate in paragraph.Descendants<AlternateContent>().ToList())
        {
            TextBoxNav.MirrorFallback(alternate);
        }
    }

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

    private static string PlainTextOf(List<Node> content) =>
        string.Join("\n", content.Select(TextOfNode));

    private static string TextOfNode(Node node) =>
        node.Text ?? string.Concat((node.Content ?? []).Select(TextOfNode));

    /// <summary>
    /// Pelos filhos diretos: o texto de uma caixa também é <c>w:t</c>, e em
    /// profundidade rejeitaria justamente os runs a copiar.
    /// </summary>
    internal static bool IsAnchoredOnly(Run run) =>
        run.Descendants<WordDrawing.Anchor>().Any() && !run.Elements<Text>().Any();

    private Paragraph WriteParagraph(Node node, ListPlacement? list, Paragraph? original)
    {
        var paragraph = new Paragraph();

        // O `w:pPr` parte do original (ParagraphFormat): estilo, espaçamento, marca e `w:sectPr` sobrevivem.
        if (_format.Build(node, list, original) is { } properties)
        {
            paragraph.ParagraphProperties = properties;
        }

        foreach (var mark in _references ? [] : Bookmarks(original, leading: true))
        {
            paragraph.AppendChild(mark.CloneNode(true));
        }

        _paragraphRun = _styles.Resolve(paragraph.ParagraphProperties).Run;

        // As imagens que já estavam voltam com o desenho original — ver ImageWriter.Reuse.
        var images = ImageWriter.FlowingImagesOf(original);
        foreach (var child in node.Content ?? [])
        {
            // Ver Revisions.Wrap.
            var written = WriteInline(child, images);
            foreach (var element in Revisions.HasRevision(child) ? Revisions.Wrap([.. written], child.Marks) : written)
            {
                paragraph.AppendChild(element);
            }
        }

        MergeAnchorLinks(paragraph);
        Revisions.MergeNeighbours(paragraph);

        // A quebra volta ao fim do parágrafo, num `w:r`, de onde o leitor a tirou.
        if (Attr.Bool(node, "breakAfter"))
        {
            paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        }

        if (Attr.Bool(node, "columnBreakAfter"))
        {
            paragraph.AppendChild(new Run(new Break { Type = BreakValues.Column }));
        }

        foreach (var mark in _references ? [] : Bookmarks(original, leading: false))
        {
            paragraph.AppendChild(mark.CloneNode(true));
        }

        return paragraph;
    }

    /// <summary>
    /// Links vizinhos para o mesmo marcador viram um só <c>w:hyperlink</c>, como o
    /// Word grava a entrada do sumário.
    /// </summary>
    private static void MergeAnchorLinks(Paragraph paragraph)
    {
        Hyperlink? previous = null;
        foreach (var child in paragraph.ChildElements.ToList())
        {
            if (child is Hyperlink link && link.Id is null && link.Anchor?.Value is { } anchor &&
                previous?.Anchor?.Value == anchor && previous.Id is null)
            {
                foreach (var inner in link.ChildElements.ToList())
                {
                    inner.Remove();
                    previous.AppendChild(inner);
                }

                link.Remove();
                continue;
            }

            previous = child as Hyperlink;
        }
    }

    /// <summary>
    /// Os marcadores do original, que o modelo não representa: sem eles, quem os cita
    /// apontaria para o vazio. Pelo lado em que estavam, para o trecho marcado não
    /// encolher; a posição no meio da frase é o que se perde.
    /// </summary>
    /// <param name="leading">Os que vinham antes de qualquer conteúdo.</param>
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
                foreach (var element in WriteTextRun(node)) yield return element;
                break;

            case "hardBreak":
                yield return new Run(new Break());
                break;

            case "pageBreak":
                yield return new Run(new Break { Type = BreakValues.Page });
                break;

            // O id e o nome do arquivo, ou os que o editor deu ao marcador novo.
            case "bookmarkStart":
                yield return new BookmarkStart
                {
                    Id = Attr.String(node, "bid") ?? "0",
                    Name = Attr.String(node, "name") ?? string.Empty,
                };
                break;

            case "bookmarkEnd":
                yield return new BookmarkEnd { Id = Attr.String(node, "bid") ?? "0" };
                break;

            // As pontas das respostas vão junto: no Word a conversa inteira abraça o trecho.
            case "commentStart":
                _comments ??= new CommentAnchors(_part);
                foreach (var id in _comments.Thread(Attr.String(node, "cid") ?? "0"))
                {
                    _comments.Start(id);
                    yield return new CommentRangeStart { Id = id };
                }

                break;

            case "commentEnd":
                _comments ??= new CommentAnchors(_part);
                foreach (var id in _comments.Thread(Attr.String(node, "cid") ?? "0"))
                {
                    if (_comments.Ranged(id)) yield return new CommentRangeEnd { Id = id };
                    yield return _comments.Reference(id);
                }

                break;

            case "field":
                foreach (var element in WriteField(node)) yield return element;
                break;

            // O corpo vai para a parte das notas (NotesWriter); aqui fica o run com o id.
            case "noteRef":
                yield return WriteNoteReference(node);
                break;

            // O OMML como veio; a nova ou editada o tira do MathML (OmmlMath.ToOmml).
            case "math":
                if (MathOf(node) is { } math) yield return math;
                else _inventory.NoteLoss("equação que não pôde ser gravada");
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

    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>O MathML não guarda cor nem fonte das fichas: a formatação de dentro se perde, e isso se diz uma vez.</summary>
    private void NoteEditedMathFormatting(Node paragraph, OpenXmlElement? original)
    {
        if (original is null || paragraph.Content is not { } content) return;
        var maths = content.Where(child => child.Type == "math").ToList();
        if (!maths.Any(math => Attr.String(math, "omml") is not { Length: > 0 })) return;

        var kept = maths.Select(math => Attr.String(math, "omml")).OfType<string>().ToHashSet(StringComparer.Ordinal);
        var equations = original.Descendants()
            .Where(element => element is DocumentFormat.OpenXml.Math.Paragraph ||
                              (element is DocumentFormat.OpenXml.Math.OfficeMath &&
                               element.Ancestors<DocumentFormat.OpenXml.Math.Paragraph>().FirstOrDefault() is null));
        if (equations.Any(equation => FormattedInside(equation) && !kept.Contains(equation.OuterXml)))
        {
            _inventory.NoteLoss("formatação dentro da equação editada");
        }
    }

    private static bool FormattedInside(OpenXmlElement equation) =>
        equation.Descendants<RunProperties>().Any(properties => properties.ChildElements.Any(child => child switch
        {
            RunFonts fonts => (fonts.Ascii?.Value ?? MathFontName) != MathFontName || (fonts.HighAnsi?.Value ?? MathFontName) != MathFontName,
            Italic or ItalicComplexScript or Languages or NoProof => false,
            _ => true,
        }));

    private const string MathFontName = "Cambria Math";

    /// <summary>Ou nulo.</summary>
    internal static OpenXmlElement? MathOf(Node node)
    {
        var omml = Attr.String(node, "omml") is { Length: > 0 } kept
            ? kept
            : Attr.String(node, "mathml") is { Length: > 0 } mathMl
                ? OmmlMath.ToOmml(mathMl, Attr.Bool(node, "display"), Attr.String(node, "jc"))
                : null;
        if (omml is null) return null;
        try
        {
            var root = System.Xml.Linq.XElement.Parse(omml);
            if (root.Name.Namespace != OmmlMath.M) return null;
            OpenXmlElement? math = root.Name.LocalName switch
            {
                "oMathPara" => new DocumentFormat.OpenXml.Math.Paragraph(omml),
                "oMath" => new DocumentFormat.OpenXml.Math.OfficeMath(omml),
                _ => null,
            };

            // A declaração do `w:` da raiz basta: a equação sai escrita como entrou.
            foreach (var element in math?.Descendants() ?? [])
            {
                foreach (var (prefix, uri) in element.NamespaceDeclarations.ToList())
                {
                    if ((prefix, uri) is ("w", WordprocessingNamespace) or ("m", MathNamespace))
                    {
                        element.RemoveNamespaceDeclaration(prefix);
                    }
                }
            }

            return math;
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }
    }

    /// <summary>O <c>w:rPr</c> do run original, ou o do Word na nota nova; a marca própria vem logo depois.</summary>
    private Run WriteNoteReference(Node node)
    {
        var endnote = Attr.String(node, "kind") == NotesWriter.Endnote;
        var id = long.TryParse(Attr.String(node, "nid"), System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var number) ? number : 0L;
        _noteRuns ??= NotesWriter.ReferenceRunsOf(_part);

        var run = new Run();
        var address = NotesWriter.Address(endnote, id.ToString(System.Globalization.CultureInfo.InvariantCulture));
        run.RunProperties = _noteRuns.TryGetValue(address, out var original) && original.RunProperties is { } kept
            ? (RunProperties)kept.CloneNode(true)
            : NotesWriter.ReferenceProperties(_part, endnote);

        FootnoteEndnoteReferenceType reference = endnote ? new EndnoteReference() : new FootnoteReference();
        reference.Id = id;
        var mark = Attr.String(node, "mark");
        if (mark is { Length: > 0 }) reference.CustomMarkFollows = true;
        run.AppendChild(reference);
        if (mark is { Length: > 0 })
        {
            foreach (var piece in XmlText.Of(mark)) run.AppendChild(piece);
        }

        return run;
    }

    private (RunProperties Properties, string? Hyperlink) FormatOf(Node node)
    {
        var properties = new RunProperties();
        string? hyperlink = null;

        foreach (var mark in node.Marks ?? [])
        {
            switch (mark.Type)
            {
                // `off` desliga o que o estilo do parágrafo liga — ver RunReader.Off.
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

                // Uma propriedade só, `w:vertAlign`, de valores que se excluem, como no editor.
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

                // Ela embrulha o run — ver WriteParagraph.
                case Revisions.Insertion:
                case Revisions.Deletion:
                    break;

                default:
                    _inventory.NoteLoss($"formatação \"{mark.Type}\"");
                    break;
            }
        }

        DropWhatRepeatsTheStyle(properties, _paragraphRun);
        return (properties, hyperlink);
    }

    private IEnumerable<OpenXmlElement> WriteTextRun(Node node)
    {
        var run = new Run();
        var (properties, hyperlink) = FormatOf(node);
        if (properties.HasChildren) run.RunProperties = properties;

        // Tabulação é `w:tab`, quebra é `w:br`, e caractere de controle não existe no XML 1.0 (XmlText).
        var pieces = XmlText.Of(node.Text).ToList();
        if (pieces.Count == 0)
        {
            pieces.Add(new Text(string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        }

        foreach (var piece in pieces) run.AppendChild(piece);

        return Linked([run], hyperlink);
    }

    private IEnumerable<OpenXmlElement> Linked(List<Run> runs, string? hyperlink)
    {
        if (hyperlink is null) return runs;

        // O link interno: `w:anchor`, sem relacionamento; `w:history`, como o Word grava.
        if (hyperlink.StartsWith('#'))
        {
            return [new Hyperlink(runs) { Anchor = hyperlink[1..], History = true }];
        }

        Uri target;
        try
        {
            target = new Uri(hyperlink, UriKind.Absolute);
        }
        catch (UriFormatException)
        {
            _inventory.NoteLoss("endereço de link inválido");
            return runs;
        }

        var relationship = _owner.AddHyperlinkRelationship(target, true);
        return [new Hyperlink(runs) { Id = relationship.Id }];
    }

    /// <summary>Cada peça no seu run, com a formatação do nó, como o Word grava.</summary>
    private IEnumerable<OpenXmlElement> WriteField(Node node)
    {
        var (properties, hyperlink) = FormatOf(node);

        Run Piece(params OpenXmlElement[] content)
        {
            var run = new Run();
            if (properties.HasChildren) run.RunProperties = (RunProperties)properties.CloneNode(true);
            foreach (var element in content) run.AppendChild(element);
            return run;
        }

        var runs = new List<Run>
        {
            Piece(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            Piece(new FieldCode(Attr.String(node, "instr") ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }),
            Piece(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        };

        var result = XmlText.Of(Attr.String(node, "result")).ToArray();
        if (result.Length > 0) runs.Add(Piece(result));
        runs.Add(Piece(new FieldChar { FieldCharType = FieldCharValues.End }));

        return Linked(runs, hyperlink);
    }

    /// <summary>
    /// Os parágrafos com o campo <c>TOC</c> em volta, no controle de conteúdo se havia um.
    /// O título volta com o <c>w:pPr</c> original; as entradas não, porque "Atualizar
    /// sumário" as refaz. A entrada sem parada ganha a do Word, à direita com pontinhos.
    /// </summary>
    private IEnumerable<OpenXmlElement> WriteTableOfContents(Node node, OpenXmlElement? original)
    {
        var children = node.Content ?? [];
        if (children.Count == 0) children = [Node.Of("paragraph")];

        var head = Math.Clamp(Attr.Int(node, "head") ?? 0, 0, children.Count - 1);
        var originals = original switch
        {
            SdtBlock block => block.SdtContentBlock?.Elements<Paragraph>().ToList() ?? [],
            Paragraph first => [first],
            _ => [],
        };

        var paragraphs = new List<Paragraph>();
        for (var index = 0; index < children.Count; index++)
        {
            var source = index < head ? originals.ElementAtOrDefault(index) : null;
            foreach (var paragraph in Write(children[index], null, source).OfType<Paragraph>())
            {
                if (index >= head) WithPageTab(paragraph);
                paragraphs.Add(paragraph);
            }
        }

        var opening = paragraphs[Math.Min(head, paragraphs.Count - 1)];
        var at = opening.ParagraphProperties is null ? 0 : 1;
        opening.InsertAt(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), at);
        opening.InsertAt(
            new Run(new FieldCode(Attr.String(node, "instr") ?? " TOC \\o \"1-3\" \\h \\z \\u ")
            {
                Space = SpaceProcessingModeValues.Preserve,
            }),
            at + 1);
        opening.InsertAt(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }), at + 2);
        paragraphs[^1].AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        if (!Attr.Bool(node, "sdt"))
        {
            foreach (var paragraph in paragraphs) yield return paragraph;
            yield break;
        }

        // As propriedades do original; o sumário novo ganha as do Word.
        var sdt = new SdtBlock();
        if (original is SdtBlock previous && previous.SdtProperties is { } kept)
        {
            sdt.AppendChild(kept.CloneNode(true));
            if (previous.SdtEndCharProperties is { } end) sdt.AppendChild(end.CloneNode(true));
        }
        else
        {
            sdt.AppendChild(new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Table of Contents" },
                    new DocPartUnique())));
        }

        sdt.AppendChild(new SdtContentBlock(paragraphs));
        yield return sdt;
    }

    private void WithPageTab(Paragraph paragraph)
    {
        if (!paragraph.Descendants<FieldCode>().Any(code => code.Text.Contains("PAGEREF", StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        if (properties.Tabs is not null) return;

        var tabs = new Tabs(new TabStop
        {
            Val = TabStopValues.Right,
            Leader = TabStopLeaderCharValues.Dot,
            Position = _usableTwips,
        });

        // Na ordem do esquema: depois do estilo, da numeração e das bordas, antes do espaçamento.
        var before = properties.ChildElements.LastOrDefault(child =>
            child is ParagraphStyleId or KeepNext or KeepLines or PageBreakBefore or FrameProperties
                or WidowControl or NumberingProperties or SuppressLineNumbers or ParagraphBorders or Shading);
        if (before is null) properties.PrependChild(tabs);
        else properties.InsertAfter(tabs, before);
    }

    /// <summary>
    /// As marcas chegam achatadas, com o estilo dentro: gravadas de volta, desligariam
    /// o parágrafo do estilo. Só o que coincide sai; marca ausente não vira "desligado".
    /// </summary>
    private static void DropWhatRepeatsTheStyle(RunProperties properties, RunProperties style)
    {
        // Ligado ou desligado, o que coincide sai.
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

        // `backgroundColor` no editor é o mesmo `w:shd` do realce no arquivo.
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
            // `w:sz` é em meios-pontos, e a medida pode chegar em pixels, como o CSS a escreve.
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

        // A entrelinha é do parágrafo (ParagraphFormat): num trecho não tem para onde ir.
        if (Attr.MarkString(mark, "lineHeight") is { } lineHeight)
        {
            _inventory.NoteLoss($"entrelinha de um trecho de texto (\"{lineHeight}\")");
        }
    }

    /// <summary>
    /// Sempre hexadecimal de seis dígitos: <c>rgb(...)</c> faria o Word declarar o
    /// documento danificado, e um nome de cor sairia preto.
    /// </summary>
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

    /// <summary>
    /// As respostas de cada conversa, os comentários com trecho (inventar um
    /// <c>w:commentRangeEnd</c> no de ponto seria errado) e o run de cada referência.
    /// </summary>
    private sealed class CommentAnchors
    {
        private readonly Dictionary<string, List<string>> _replies;
        private readonly HashSet<string> _ranged;
        private readonly Dictionary<string, Run> _references = new(StringComparer.Ordinal);

        /// <summary>O comentário criado no editor não está no original: é o começo dele que diz que tem trecho.</summary>
        private readonly HashSet<string> _started = new(StringComparer.Ordinal);

        public CommentAnchors(MainDocumentPart part)
        {
            _replies = CommentsReader.RepliesOf(part);
            // O parágrafo da nota também leva âncora.
            var roots = CommentsWriter.AnchorRoots(part);
            _ranged = roots.SelectMany(root => root.Descendants<CommentRangeStart>())
                .Select(start => start.Id?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal);
            foreach (var run in roots.SelectMany(root => root.Descendants<Run>()))
            {
                if (BodyReader.ReferenceOnly(run) is { } id) _references.TryAdd(id, run);
            }
        }

        public IEnumerable<string> Thread(string id) => [id, .. _replies.GetValueOrDefault(id) ?? []];

        public void Start(string id) => _started.Add(id);

        public bool Ranged(string id) => _ranged.Contains(id) || _started.Contains(id);

        public Run Reference(string id) =>
            _references.TryGetValue(id, out var original)
                ? (Run)original.CloneNode(true)
                : new Run(
                    new RunProperties(new RunStyle { Val = "CommentReference" }),
                    new CommentReference { Id = id });
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Tests;

/// <summary>
/// Documentos de teste construídos em código, não versionados como binário.
/// </summary>
/// <remarks>
/// O corpus real da empresa tem marca de cliente e capturas de sistemas
/// internos — não entra no repositório
/// (docs/01-corpus-docx.md). Estes fixtures reproduzem as **estruturas**
/// catalogadas lá com conteúdo inventado.
///
/// Construir em código em vez de guardar `.docx` tem uma vantagem que não é
/// óbvia: o que o fixture contém fica legível na revisão. Um binário no git é
/// uma caixa preta que ninguém confere.
/// </remarks>
public static class Fixtures
{
    /// <summary>Documento simples com três parágrafos.</summary>
    public static byte[] Simple() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Primeiro parágrafo.", style: "Heading1"));
        body.AppendChild(Paragraph("Segundo parágrafo, com texto comum."));
        body.AppendChild(Paragraph("Terceiro parágrafo."));
    });

    /// <summary>
    /// Documento com um comentário ancorado no **segundo** parágrafo.
    /// </summary>
    public static byte[] WithComment() => Build((body, part) =>
    {
        var comments = part.AddNewPart<WordprocessingCommentsPart>();
        comments.Comments = new Comments(
            new Comment(new Paragraph(new Run(new Text("Revisar esta frase."))))
            {
                Id = "1",
                Author = "Revisor",
                Date = System.Xml.XmlConvert.ToDateTime(
                    "2026-01-01T00:00:00Z", System.Xml.XmlDateTimeSerializationMode.Utc),
            });

        body.AppendChild(Paragraph("Parágrafo sem comentário."));

        var commented = Paragraph("Parágrafo comentado.");
        commented.PrependChild(new CommentRangeStart { Id = "1" });
        commented.AppendChild(new CommentRangeEnd { Id = "1" });
        commented.AppendChild(new Run(new CommentReference { Id = "1" }));
        body.AppendChild(commented);

        body.AppendChild(Paragraph("Outro parágrafo sem comentário."));
    });

    /// <summary>Documento com controle de alterações no segundo parágrafo.</summary>
    public static byte[] WithTrackedChanges() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Parágrafo intocado."));

        var revised = new Paragraph();
        revised.AppendChild(new Run(new Text("Texto original ") { Space = SpaceProcessingModeValues.Preserve }));
        revised.AppendChild(new InsertedRun(
            new Run(new Text("e um acréscimo") { Space = SpaceProcessingModeValues.Preserve }))
        {
            Id = "10",
            Author = "Revisor",
        });
        revised.AppendChild(new DeletedRun(
            new Run(new DeletedText(" trecho removido") { Space = SpaceProcessingModeValues.Preserve }))
        {
            Id = "11",
            Author = "Revisor",
        });
        body.AppendChild(revised);

        body.AppendChild(Paragraph("Outro parágrafo intocado."));
    });

    /// <summary>
    /// Imagem ancorada e centralizada, do jeito que o LibreOffice grava.
    /// Ver docs/01-corpus-docx.md, Descoberta 3.
    /// </summary>
    public static byte[] WithAnchoredImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(Paragraph("Antes da imagem."));
        body.AppendChild(new Paragraph(new Run(AnchoredDrawing(part.GetIdOfPart(image)))));
        body.AppendChild(Paragraph("Depois da imagem."));
    });

    /// <summary>
    /// Quebra de página **dentro** do parágrafo, no fim de um `w:r`.
    /// </summary>
    /// <remarks>
    /// É como o Word grava "daqui para frente é outra página" sem fechar o
    /// parágrafo. Distinta da quebra que ocupa um parágrafo só dela.
    /// </remarks>
    public static byte[] WithBreakInsideParagraph() => Build((body, _) =>
    {
        var comQuebra = new Paragraph();
        comQuebra.AppendChild(new Run(new Text("Fim da primeira página.")));
        comQuebra.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        body.AppendChild(comQuebra);
        body.AppendChild(Paragraph("Começo da segunda."));
    });

    /// <summary>
    /// Três cabeçalhos declarados, com o `first` **antes** do `default` no XML.
    /// </summary>
    /// <remarks>
    /// A ordem é escolhida para enganar: o leitor antigo percorria as
    /// referências na ordem de gravação e ficava com a primeira não vazia, então
    /// o cabeçalho da capa apareceria em todas as páginas. Qual aparecia
    /// dependia de como o Word gravou, e não do que o documento diz.
    /// </remarks>
    public static byte[] WithFirstPageHeader(bool titlePage) => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            if (titlePage) section.AppendChild(new TitlePage());
            section.AppendChild(new HeaderReference { Type = HeaderFooterValues.First, Id = Header(part, "Capa") });
            section.AppendChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = Header(part, "Miolo") });
        });

    /// <summary>Um cabeçalho novo com uma linha de texto; devolve o `r:id`.</summary>
    private static string Header(MainDocumentPart part, string text)
    {
        var header = part.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text(text))));
        header.Header.Save();
        return part.GetIdOfPart(header);
    }

    /// <summary>
    /// Duas caixas de texto ancoradas no **mesmo** parágrafo, como o Word grava
    /// a capa de um modelo de manual: título e subtítulo em caixas separadas.
    /// </summary>
    /// <remarks>
    /// Cada caixa vem duas vezes no arquivo — `mc:Choice` em DrawingML e
    /// `mc:Fallback` no VML antigo, com o mesmo texto dentro. É essa duplicação
    /// que faz o leitor precisar escolher um ramo: percorrer os dois escreveria
    /// cada título duas vezes na tela.
    /// </remarks>
    public static byte[] WithTextBoxes() => Build((body, _) =>
    {
        var cover = new Paragraph();
        cover.AppendChild(new Run(TextBoxShape("Título do manual")));
        cover.AppendChild(new Run(TextBoxShape("Subtítulo do manual")));
        body.AppendChild(cover);
        body.AppendChild(Paragraph("Primeiro parágrafo do corpo."));
    });

    /// <summary>
    /// Imagem ancorada girada um quarto de volta — a marca vertical que corre
    /// pela lateral da capa.
    /// </summary>
    /// <remarks>
    /// `wp:extent` mede a imagem deitada: 28,58 cm de comprido por 8,01 cm de
    /// altura. Girada, o que ocupa a largura da página são os 8,01 cm.
    /// </remarks>
    public static byte[] WithRotatedImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(new Paragraph(new Run(
            AnchoredDrawing(part.GetIdOfPart(image), cx: 10287325, cy: 2885145, rotation: 16200000))));
    });

    /// <summary>
    /// Sete seções `continuous` com geometria idêntica — o artefato do
    /// LibreOffice descrito na Descoberta 5 do corpus.
    /// </summary>
    public static byte[] WithIdenticalSections() => Build((body, _) =>
    {
        for (var i = 1; i <= 6; i++)
        {
            var paragraph = Paragraph($"Trecho {i}.");
            // Mesma geometria da seção final que `Build` acrescenta — é isto
            // que caracteriza o artefato: sete seções dizendo a mesma coisa.
            paragraph.ParagraphProperties = new ParagraphProperties(
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.Continuous },
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U }));
            body.AppendChild(paragraph);
        }

        body.AppendChild(Paragraph("Fim."));
    });

    /// <summary>
    /// Duas seções que **divergem** de verdade: a primeira em paisagem.
    /// Aqui o autor quis duas configurações, e o nosso modelo de uma só perde.
    /// </summary>
    public static byte[] WithDivergentSections() => Build((body, _) =>
    {
        var paragraph = Paragraph("Trecho em paisagem.");
        paragraph.ParagraphProperties = new ParagraphProperties(
            new SectionProperties(
                new SectionType { Val = SectionMarkValues.NextPage },
                new PageSize { Width = 16838U, Height = 11906U, Orient = PageOrientationValues.Landscape },
                new PageMargin { Top = 720, Bottom = 720, Left = 720U, Right = 720U }));
        body.AppendChild(paragraph);

        body.AppendChild(Paragraph("Trecho em retrato."));
    });

    /// <summary>
    /// Documento cuja formatação mora nos **estilos**, como o corpus real.
    /// </summary>
    /// <remarks>
    /// `Faixa` reproduz o `Heading1` do corpus: fundo vermelho, texto branco,
    /// Arial 10 pt, centralizado. Nada disso está no parágrafo.
    /// `Corpo` herda de `Base` para exercitar a cadeia de `basedOn`.
    /// </remarks>
    public static byte[] WithStyles() => Build((body, part) =>
    {
        var styles = part.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" },
                    new FontSize { Val = "22" }))),

            new Style(
                new StyleName { Val = "Faixa" },
                new StyleParagraphProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "943634" },
                    new Justification { Val = JustificationValues.Center }),
                new StyleRunProperties(
                    new RunFonts { Ascii = "Arial" },
                    new Bold(),
                    new Color { Val = "FFFFFF" },
                    new FontSize { Val = "20" }))
            { Type = StyleValues.Paragraph, StyleId = "Faixa" },

            new Style(
                new StyleName { Val = "Base" },
                new StyleRunProperties(new RunFonts { Ascii = "Arial" }, new FontSize { Val = "20" }))
            { Type = StyleValues.Paragraph, StyleId = "Base" },

            new Style(
                new StyleName { Val = "Corpo" },
                new BasedOn { Val = "Base" },
                new StyleParagraphProperties(new Justification { Val = JustificationValues.Both }))
            { Type = StyleValues.Paragraph, StyleId = "Corpo" });

        body.AppendChild(Paragraph("Informações Gerais", style: "Faixa"));
        body.AppendChild(Paragraph("Texto do corpo.", style: "Corpo"));

        // Formatação direta tem de vencer o estilo.
        var overridden = Paragraph("Alinhado à direita.", style: "Corpo");
        overridden.ParagraphProperties!.AppendChild(new Justification { Val = JustificationValues.Right });
        body.AppendChild(overridden);

        // `w:rPr` dentro de `w:pPr` formata a marca de parágrafo, não os runs.
        var marked = Paragraph("Não deve ficar em negrito.", style: "Corpo");
        marked.ParagraphProperties!.AppendChild(new ParagraphMarkRunProperties(new Bold()));
        body.AppendChild(marked);
    });

    /// <summary>Um parágrafo que pede para ficar com o seguinte, outro que não.</summary>
    public static byte[] WithKeepNext() => Build((body, _) =>
    {
        var kept = Paragraph("Rótulo da imagem:");
        kept.ParagraphProperties = new ParagraphProperties(new KeepNext());
        body.AppendChild(kept);

        body.AppendChild(Paragraph("Solto no meio do texto."));
    });

    /// <summary>
    /// O idioma do corpus: alinhado à esquerda, mas centralizado por tabulação.
    /// </summary>
    /// <remarks>
    /// O autor põe `w:jc` em `left`, define uma parada de tabulação
    /// centralizada no meio da coluna e usa `Tab` para chegar até ela.
    /// </remarks>
    public static byte[] WithTabCentering() => Build((body, _) =>
    {
        var centered = new Paragraph();
        centered.ParagraphProperties = new ParagraphProperties(
            new Tabs(new TabStop { Val = TabStopValues.Center, Position = 4153 }),
            new Justification { Val = JustificationValues.Left });
        centered.AppendChild(new Run(new TabChar(), new TabChar(), new Text("Centralizado por tabulação")));
        body.AppendChild(centered);

        // Tabulação no meio da linha não é posicionamento de parágrafo.
        var inline = new Paragraph();
        inline.ParagraphProperties = new ParagraphProperties(
            new Tabs(new TabStop { Val = TabStopValues.Center, Position = 4153 }),
            new Justification { Val = JustificationValues.Left });
        inline.AppendChild(new Run(new Text("Esquerda"), new TabChar(), new Text("Meio")));
        body.AppendChild(inline);
    });

    /// <summary>Estilo que herda de si mesmo — não pode travar o leitor.</summary>
    public static byte[] WithCircularStyle() => Build((body, part) =>
    {
        var styles = part.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(
            new Style(new StyleName { Val = "A" }, new BasedOn { Val = "B" })
            { Type = StyleValues.Paragraph, StyleId = "A" },
            new Style(new StyleName { Val = "B" }, new BasedOn { Val = "A" })
            { Type = StyleValues.Paragraph, StyleId = "B" });

        body.AppendChild(Paragraph("Texto.", style: "A"));
    });

    /// <summary>Versalete e maiúsculas — 45 ocorrências no corpus.</summary>
    public static byte[] WithSmallCaps() => Build((body, _) =>
    {
        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(new RunProperties(new SmallCaps()), new Text("versalete")));
        paragraph.AppendChild(new Run(new RunProperties(new Caps()), new Text("maiúsculas")));
        body.AppendChild(paragraph);
    });

    public static byte[] WithTable() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Antes da tabela."));

        var table = new Table(new TableProperties(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 })));

        foreach (var row in new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } })
        {
            var tableRow = new TableRow();
            foreach (var cell in row)
            {
                tableRow.AppendChild(new TableCell(Paragraph(cell)));
            }

            table.AppendChild(tableRow);
        }

        body.AppendChild(table);
        body.AppendChild(Paragraph("Depois da tabela."));
    });

    /// <summary>Lista com marcador, com a numeração declarada de verdade.</summary>
    public static byte[] WithBulletList() => Build((body, part) =>
    {
        var numbering = part.AddNewPart<NumberingDefinitionsPart>();
        numbering.Numbering = new Numbering(
            new AbstractNum(new Level(new NumberingFormat { Val = NumberFormatValues.Bullet })
            {
                LevelIndex = 0,
            })
            { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });

        body.AppendChild(Paragraph("Introdução."));

        foreach (var item in new[] { "Primeiro item", "Segundo item" })
        {
            var paragraph = Paragraph(item);
            paragraph.ParagraphProperties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            body.AppendChild(paragraph);
        }

        body.AppendChild(Paragraph("Conclusão."));
    });

    // --- construção ---------------------------------------------------------

    private static Paragraph Paragraph(string text, string? style = null)
    {
        var paragraph = new Paragraph();
        if (style is not null)
        {
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = style });
        }

        paragraph.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static byte[] Build(Action<Body, MainDocumentPart> fill) => Build(fill, null);

    /// <param name="decorate">
    /// Recebe o `w:sectPr` já montado, para o fixture acrescentar referências de
    /// cabeçalho ou interruptores de seção.
    /// </param>
    private static byte[] Build(
        Action<Body, MainDocumentPart> fill,
        Action<SectionProperties, MainDocumentPart>? decorate)
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var part = document.AddMainDocumentPart();
            var body = new Body();

            fill(body, part);

            var section = new SectionProperties(
                new PageSize { Width = 11906U, Height = 16838U },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U });
            decorate?.Invoke(section, part);
            body.AppendChild(section);

            part.Document = new Document(body);
            part.Document.Save();
        }

        return buffer.ToArray();
    }

    /// <summary>
    /// Uma caixa de texto do jeito que o Word grava: o mesmo texto em
    /// DrawingML e no VML de reserva.
    /// </summary>
    private static OpenXmlElement TextBoxShape(string text)
    {
        var xml = $"""
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <mc:Choice Requires="wps">
                <w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                           xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="3"
                             behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
                    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                    <wp:extent cx="3800475" cy="2019300"/>
                    <wp:wrapNone/>
                    <wp:docPr id="7" name="Caixa de Texto"/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr txBox="1"/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="3800475" cy="2019300"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr rot="0" vert="horz" wrap="square"/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:anchor>
                </w:drawing>
              </mc:Choice>
              <mc:Fallback>
                <w:pict xmlns:v="urn:schemas-microsoft-com:vml">
                  <v:shape id="Caixa_{Guid.NewGuid():N}" type="#_x0000_t202"
                           style="position:absolute;width:299.25pt;height:159pt">
                    <v:textbox>
                      <w:txbxContent>
                        <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
                      </w:txbxContent>
                    </v:textbox>
                  </v:shape>
                </w:pict>
              </mc:Fallback>
            </mc:AlternateContent>
            """;

        // Construído a partir do XML completo: `InnerXml` receberia só os
        // ramos e perderia o elemento que os envolve.
        return new AlternateContent(xml);
    }

    /// <summary>Imagem **no fluxo** (`wp:inline`): ocupa lugar na linha.</summary>
    public static byte[] WithInlineImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(new Paragraph(new Run(InlineDrawing(part.GetIdOfPart(image)))));
    });

    private static OpenXmlElement InlineDrawing(string relationshipId)
    {
        var xml = $"""
            <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       distT="0" distB="0" distL="0" distR="0">
              <wp:extent cx="5274000" cy="2637000"/>
              <wp:docPr id="1" name="Imagem 1"/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id="1" name="Imagem 1"/><pic:cNvPicPr/></pic:nvPicPr>
                    <pic:blipFill><a:blip r:embed="{relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr>
                      <a:xfrm><a:off x="0" y="0"/><a:ext cx="5274000" cy="2637000"/></a:xfrm>
                      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    </pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
            """;

        var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
        drawing.InnerXml = xml;
        return drawing;
    }

    private static OpenXmlElement AnchoredDrawing(
        string relationshipId,
        long cx = 5274000,
        long cy = 2637000,
        int rotation = 0)
    {
        const string Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        var xml = $"""
            <wp:anchor xmlns:wp="{Wp}"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2"
                       behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">
              <wp:simplePos x="0" y="0"/>
              <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
              <wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV>
              <wp:extent cx="{cx}" cy="{cy}"/>
              <wp:wrapSquare wrapText="bothSides"/>
              <wp:docPr id="1" name="Imagem 1"/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr>
                      <pic:cNvPr id="1" name="Imagem 1"/>
                      <pic:cNvPicPr/>
                    </pic:nvPicPr>
                    <pic:blipFill>
                      <a:blip r:embed="{relationshipId}"/>
                      <a:stretch><a:fillRect/></a:stretch>
                    </pic:blipFill>
                    <pic:spPr>
                      <a:xfrm rot="{rotation}"><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    </pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:anchor>
            """;

        var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
        drawing.InnerXml = xml;
        return drawing;
    }

    /// <summary>PNG 1×1 transparente, o menor válido.</summary>
    private static byte[] TinyPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
}

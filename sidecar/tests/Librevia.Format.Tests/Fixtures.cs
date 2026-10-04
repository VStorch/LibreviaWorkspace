using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Tests;

/// <summary>
/// Documentos de teste montados em código, legíveis na revisão. Reproduzem as
/// estruturas do corpus, que não entra no repositório, com conteúdo inventado.
/// </summary>
public static class Fixtures
{
    /// <summary>Documento simples com três parágrafos.</summary>
    public static byte[] Simple() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Primeiro parágrafo.", style: "Heading1"));
        body.AppendChild(Paragraph("Segundo parágrafo, com texto comum."));
        body.AppendChild(Paragraph("Terceiro parágrafo."));
    });

    /// <summary>Um comentário ancorado no segundo parágrafo.</summary>
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

    /// <summary>
    /// Uma conversa, um comentário resolvido e um de ponto, como o Word grava: a
    /// resposta abraça o trecho do pai, e <c>commentsExtended.xml</c> liga as duas pelo
    /// <c>w14:paraId</c>. O de ponto só tem a referência.
    /// </summary>
    public static byte[] WithCommentThread() => Build((body, part) =>
    {
        var date = System.Xml.XmlConvert.ToDateTime(
            "2026-03-02T10:00:00Z", System.Xml.XmlDateTimeSerializationMode.Utc);
        Comment Of(string id, string author, string paraId, params string[] lines) =>
            new(lines.Select(line => new Paragraph(new Run(new Text(line))) { ParagraphId = paraId }))
            {
                Id = id,
                Author = author,
                Initials = author[..1],
                Date = date,
            };

        var comments = part.AddNewPart<WordprocessingCommentsPart>();
        comments.Comments = new Comments(
            Of("0", "Ana", "10000000", "Conferir o valor."),
            Of("1", "Bruno", "10000001", "Conferido."),
            Of("2", "Ana", "10000002", "Já resolvido."),
            Of("3", "Carla", "10000003", "Ponto sem trecho."));
        var rich = new Paragraph(new Run(new RunProperties(new Bold()), new Text("negrito")));
        comments.Comments.AppendChild(new Comment(rich) { Id = "4", Author = "Dora", Date = date });

        var extended = part.AddNewPart<WordprocessingCommentsExPart>();
        extended.CommentsEx = new DocumentFormat.OpenXml.Office2013.Word.CommentsEx(
            new DocumentFormat.OpenXml.Office2013.Word.CommentEx { ParaId = "10000000", Done = false },
            new DocumentFormat.OpenXml.Office2013.Word.CommentEx
            {
                ParaId = "10000001", ParaIdParent = "10000000", Done = false,
            },
            new DocumentFormat.OpenXml.Office2013.Word.CommentEx { ParaId = "10000002", Done = true },
            new DocumentFormat.OpenXml.Office2013.Word.CommentEx { ParaId = "10000003", Done = false });

        body.AppendChild(Paragraph("Parágrafo sem comentário."));

        var thread = new Paragraph(
            new Run(new Text("O valor ") { Space = SpaceProcessingModeValues.Preserve }),
            new CommentRangeStart { Id = "0" },
            new CommentRangeStart { Id = "1" },
            new Run(new Text("doze mil")),
            new CommentRangeEnd { Id = "0" },
            new Run(new RunProperties(new RunStyle { Val = "CommentReference" }), new CommentReference { Id = "0" }),
            new CommentRangeEnd { Id = "1" },
            new Run(new RunProperties(new RunStyle { Val = "CommentReference" }), new CommentReference { Id = "1" }),
            new Run(new Text(" consta da ata.") { Space = SpaceProcessingModeValues.Preserve }));
        body.AppendChild(thread);

        var resolved = Paragraph("Parágrafo resolvido.");
        resolved.PrependChild(new CommentRangeStart { Id = "2" });
        resolved.AppendChild(new CommentRangeEnd { Id = "2" });
        resolved.AppendChild(new Run(new CommentReference { Id = "2" }));
        body.AppendChild(resolved);

        var point = Paragraph("Parágrafo com ponto.");
        point.AppendChild(new Run(new CommentReference { Id = "3" }));
        point.AppendChild(new CommentRangeStart { Id = "4" });
        point.AppendChild(new CommentRangeEnd { Id = "4" });
        point.AppendChild(new Run(new CommentReference { Id = "4" }));
        body.AppendChild(point);
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

    /// <summary>Imagem ancorada e centralizada, como o LibreOffice grava.</summary>
    public static byte[] WithAnchoredImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(Paragraph("Antes da imagem."));
        // Com posição de verdade: fora da coluna, e o texto passa por baixo.
        body.AppendChild(new Paragraph(new Run(
            AnchoredDrawing(part.GetIdOfPart(image), placed: true))));
        body.AppendChild(Paragraph("Depois da imagem."));
    });

    /// <summary>
    /// Logotipo ancorado à direita no cabeçalho. A posição mora na âncora, e o
    /// <c>a:off</c> do desenho é zero, como em todo desenho de peça única.
    /// </summary>
    public static byte[] WithAnchoredHeaderLogo(long horizontalOffsetEmus) => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo.")),
        (section, part) =>
        {
            var header = part.AddNewPart<HeaderPart>();
            var image = header.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(TinyPng()))
            {
                image.FeedData(stream);
            }

            var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
            drawing.InnerXml = $"""
                <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                           xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                           distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2"
                           behindDoc="1" locked="0" layoutInCell="0" allowOverlap="1">
                  <wp:simplePos x="0" y="0"/>
                  <wp:positionH relativeFrom="column"><wp:posOffset>{horizontalOffsetEmus}</wp:posOffset></wp:positionH>
                  <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                  <wp:extent cx="1332865" cy="314325"/>
                  <wp:wrapNone/>
                  <wp:docPr id="9" name="Logotipo"/>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <pic:pic>
                        <pic:nvPicPr><pic:cNvPr id="9" name="Logotipo"/><pic:cNvPicPr/></pic:nvPicPr>
                        <pic:blipFill><a:blip r:embed="{header.GetIdOfPart(image)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                        <pic:spPr>
                          <a:xfrm><a:off x="0" y="0"/><a:ext cx="1332865" cy="314325"/></a:xfrm>
                          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        </pic:spPr>
                      </pic:pic>
                    </a:graphicData>
                  </a:graphic>
                </wp:anchor>
                """;

            header.Header = new Header(new Paragraph(new Run(drawing)));
            header.Header.Save();
            section.AppendChild(new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = part.GetIdOfPart(header),
            });
        });

    /// <summary>
    /// Quebra de página no fim de um <c>w:r</c>, dentro do parágrafo, distinta da que
    /// ocupa um parágrafo só dela.
    /// </summary>
    public static byte[] WithBreakInsideParagraph() => Build((body, _) =>
    {
        var comQuebra = new Paragraph();
        comQuebra.AppendChild(new Run(new Text("Fim da primeira página.")));
        comQuebra.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        body.AppendChild(comQuebra);
        body.AppendChild(Paragraph("Começo da segunda."));
    });

    /// <summary>Rodapé de três parágrafos centralizados, como o do modelo de manual.</summary>
    public static byte[] WithFooterOfThreeLines() => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            var footer = part.AddNewPart<FooterPart>();
            var content = new Footer();

            foreach (var text in new[] { "www.exemplo.com.br", "Documento V01 - Desenvolvido por: Fulano", "Mês/ANO" })
            {
                var paragraph = new Paragraph(new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }));
                paragraph.AppendChild(new Run(new Text(text)));
                content.AppendChild(paragraph);
            }

            footer.Footer = content;
            footer.Footer.Save();

            section.AppendChild(new FooterReference
            {
                Type = HeaderFooterValues.Default,
                Id = part.GetIdOfPart(footer),
            });
        });

    /// <summary>
    /// Rodapé com texto, tabulação e o campo <c>PAGE</c>: só o texto tem <c>w:t</c>
    /// onde escrever.
    /// </summary>
    public static byte[] WithFooterOfPageNumber() => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            var footer = part.AddNewPart<FooterPart>();
            var paragraph = new Paragraph();

            paragraph.AppendChild(new Run(new Text("Página ") { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            paragraph.AppendChild(new Run(new FieldCode(" PAGE ")));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            paragraph.AppendChild(new Run(new Text("7")));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

            footer.Footer = new Footer(paragraph);
            footer.Footer.Save();

            section.AppendChild(new FooterReference
            {
                Type = HeaderFooterValues.Default,
                Id = part.GetIdOfPart(footer),
            });
        });

    /// <summary>Quebra de página sozinha num parágrafo.</summary>
    public static byte[] WithLonePageBreak() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Primeira página."));
        body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        body.AppendChild(Paragraph("Segunda página."));
    });

    /// <summary>
    /// Três cabeçalhos, com o <c>first</c> antes do <c>default</c> no XML: o que vale é
    /// o tipo, e não a ordem de gravação.
    /// </summary>
    public static byte[] WithFirstPageHeader(bool titlePage) => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            if (titlePage) section.AppendChild(new TitlePage());
            section.AppendChild(new HeaderReference { Type = HeaderFooterValues.First, Id = Header(part, "Capa") });
            section.AppendChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = Header(part, "Miolo") });
        });

    /// <summary>
    /// Cabeçalho em grade, como o corporativo do corpus: quatro colunas e três linhas,
    /// o logotipo mesclado verticalmente na primeira e duas colunas unidas por
    /// <c>w:gridSpan</c> à direita.
    /// </summary>
    public static byte[] WithHeaderGrid() => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            var header = part.AddNewPart<HeaderPart>();
            var image = header.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(TinyPng()))
            {
                image.FeedData(stream);
            }

            var borders = new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 });

            var table = new Table(
                new TableProperties(new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, borders),
                new TableGrid(
                    new GridColumn { Width = "2000" },
                    new GridColumn { Width = "6000" },
                    new GridColumn { Width = "1000" },
                    new GridColumn { Width = "1000" }));

            table.AppendChild(new TableRow(
                LogoCell(header.GetIdOfPart(image), MergedCellValues.Restart),
                TextCell("Chamado 10001", "6000"),
                TextCell("Data de revisão", "2000", span: 2)));

            table.AppendChild(new TableRow(
                LogoCell(null, MergedCellValues.Continue),
                TextCell("Título do documento", "6000"),
                TextCell("30/07/2026", "2000", span: 2)));

            // Sem a borda de baixo, duas linhas do arquivo viram uma moldura só.
            var sem = TextCell("Página", "1000");
            sem.TableCellProperties!.AppendChild(
                new TableCellBorders(new BottomBorder { Val = BorderValues.Nil }));

            table.AppendChild(new TableRow(
                LogoCell(null, MergedCellValues.Continue),
                TextCell("Rodapé do cabeçalho", "6000"),
                sem,
                TextCell("Revisão", "1000")));

            header.Header = new Header(table, new Paragraph());
            header.Header.Save();
            section.AppendChild(new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = part.GetIdOfPart(header),
            });
        });

    private static TableCell LogoCell(string? relationshipId, MergedCellValues merge)
    {
        var properties = new TableCellProperties(
            new TableCellWidth { Width = "2000", Type = TableWidthUnitValues.Dxa },
            new VerticalMerge { Val = merge });

        var paragraph = new Paragraph(new ParagraphProperties(
            new Justification { Val = JustificationValues.Center }));

        if (relationshipId is not null)
        {
            paragraph.AppendChild(new Run(InlineDrawing(relationshipId)));
        }

        return new TableCell(properties, paragraph);
    }

    private static TableCell TextCell(string text, string width, int span = 1)
    {
        var properties = new TableCellProperties(
            new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa });
        if (span > 1) properties.AppendChild(new GridSpan { Val = span });

        return new TableCell(
            properties,
            new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
    }

    /// <summary>
    /// Fonte que a máquina pode não ter, com o tipo declarado em <c>word/fontTable.xml</c>.
    /// </summary>
    public static byte[] WithMissingFont() => Build((body, part) =>
    {
        var table = part.AddNewPart<FontTablePart>();
        table.Fonts = new Fonts(
            new Font(new FontFamily { Val = FontFamilyValues.Swiss }) { Name = "Segoe UI" },
            new Font(new FontFamily { Val = FontFamilyValues.Modern }) { Name = "Consolas" });
        table.Fonts.Save();

        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(
            new RunProperties(new RunFonts { Ascii = "Segoe UI", HighAnsi = "Segoe UI" }),
            new Text("Título da capa") { Space = SpaceProcessingModeValues.Preserve }));
        body.AppendChild(paragraph);

        // Fonte que a tabela não declara: nada a inventar, sai como está.
        var outro = new Paragraph();
        outro.AppendChild(new Run(
            new RunProperties(new RunFonts { Ascii = "Fonte Fantasma", HighAnsi = "Fonte Fantasma" }),
            new Text("Sem tipo declarado") { Space = SpaceProcessingModeValues.Preserve }));
        body.AppendChild(outro);
    });

    /// <summary>
    /// Cabeçalho que é um grupo de formas, logotipo e caixa de título: a âncora dá
    /// posição e tamanho do grupo, e <c>a:chOff</c>/<c>a:chExt</c> a régua de dentro.
    /// </summary>
    public static byte[] WithHeaderGroup() => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            var header = part.AddNewPart<HeaderPart>();
            var image = header.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(TinyPng()))
            {
                image.FeedData(stream);
            }

            var drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
            drawing.InnerXml = $"""
                <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                           xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                           xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                           xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                           xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                           distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2"
                           behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">
                  <wp:simplePos x="0" y="0"/>
                  <wp:positionH relativeFrom="page"><wp:posOffset>1143000</wp:posOffset></wp:positionH>
                  <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
                  <wp:extent cx="6371640" cy="604440"/>
                  <wp:wrapSquare wrapText="bothSides"/>
                  <wp:docPr id="9" name="Group 1"/>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
                      <wpg:wgp>
                        <wpg:cNvGrpSpPr/>
                        <wpg:grpSpPr>
                          <a:xfrm>
                            <a:off x="0" y="0"/><a:ext cx="6371640" cy="604440"/>
                            <a:chOff x="0" y="0"/><a:chExt cx="6371640" cy="604440"/>
                          </a:xfrm>
                        </wpg:grpSpPr>
                        <pic:pic>
                          <pic:nvPicPr><pic:cNvPr id="1" name="Logotipo"/><pic:cNvPicPr/></pic:nvPicPr>
                          <pic:blipFill><a:blip r:embed="{header.GetIdOfPart(image)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                          <pic:spPr>
                            <a:xfrm><a:off x="4644000" y="0"/><a:ext cx="1727640" cy="378000"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          </pic:spPr>
                        </pic:pic>
                        <wps:wsp>
                          <wps:cNvSpPr txBox="1"/>
                          <wps:spPr>
                            <a:xfrm><a:off x="1500000" y="248400"/><a:ext cx="3052800" cy="327600"/></a:xfrm>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/></w:rPr>
                                <w:t>EVIDÊNCIAS DO ROTEIRO</w:t></w:r></w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr/>
                        </wps:wsp>
                        <wps:wsp>
                          <wps:cNvSpPr/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="580000"/><a:ext cx="6371640" cy="0"/></a:xfrm>
                            <a:ln w="6480"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                          </wps:spPr>
                          <wps:bodyPr/>
                        </wps:wsp>
                      </wpg:wgp>
                    </a:graphicData>
                  </a:graphic>
                </wp:anchor>
                """;

            header.Header = new Header(new Paragraph(new Run(drawing)));
            header.Header.Save();
            section.AppendChild(new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = part.GetIdOfPart(header),
            });
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
    /// Título e subtítulo em duas caixas ancoradas no mesmo parágrafo, como a capa do
    /// modelo de manual; cada uma em <c>mc:Choice</c> e em <c>mc:Fallback</c>.
    /// </summary>
    public static byte[] WithTextBoxes() => Build((body, _) =>
    {
        var cover = new Paragraph();
        cover.AppendChild(new Run(TextBoxShape("Título do manual")));
        cover.AppendChild(new Run(TextBoxShape("Subtítulo do manual")));
        body.AppendChild(cover);
        body.AppendChild(Paragraph("Primeiro parágrafo do corpo."));
    });

    /// <summary>
    /// Três caixas: sem decoração (como as do corpus, <c>a:noFill</c> e linha zero),
    /// com uma que se desenha e com uma que não.
    /// </summary>
    public static byte[] WithDecoratedTextBoxes() => Build((body, _) =>
    {
        // Sem moldura, como as caixas do cabeçalho do corpus.
        body.AppendChild(new Paragraph(new Run(TextBoxShape(
            "Sem moldura",
            """<a:noFill/><a:ln w="0"><a:noFill/></a:ln>"""))));

        // Preenchimento e traço sólidos: o que o CSS desenha.
        body.AppendChild(new Paragraph(new Run(TextBoxShape(
            "Com moldura",
            """
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            <a:ln w="12700"><a:solidFill><a:srgbClr val="1F5FA9"/></a:solidFill></a:ln>
            """))));

        body.AppendChild(Paragraph("Corpo do documento."));
    });

    /// <summary>Caixa com preenchimento em gradiente, que não sabemos desenhar.</summary>
    public static byte[] WithGradientTextBox() => Build((body, _) =>
    {
        body.AppendChild(new Paragraph(new Run(TextBoxShape(
            "Com gradiente",
            """
            <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill>
            <a:ln w="0"><a:noFill/></a:ln>
            """))));

        body.AppendChild(Paragraph("Corpo do documento."));
    });

    /// <summary>
    /// A marca vertical da capa: imagem ancorada girada um quarto de volta. O
    /// <c>wp:extent</c> mede a imagem deitada, 28,58 × 8,01 cm.
    /// </summary>
    public static byte[] WithRotatedImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(new Paragraph(new Run(
            AnchoredDrawing(
                part.GetIdOfPart(image),
                cx: 10287325,
                cy: 2885145,
                rotation: 16200000,
                placed: true))));
    });

    /// <summary>
    /// Como o modelo de manual encerra a capa: um parágrafo só com a marca
    /// posicionada e o <c>w:br</c> da folha seguinte.
    /// </summary>
    public static byte[] WithBreakOnAnchorParagraph() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(
            AnchoredDrawing(
                part.GetIdOfPart(image),
                cx: 10287325,
                cy: 2885145,
                rotation: 16200000,
                placed: true)));
        paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));

        body.AppendChild(paragraph);
        body.AppendChild(Paragraph("Depois da capa."));
    });

    /// <summary>
    /// Texto, captura ancorada ao topo do parágrafo e mais texto, no mesmo
    /// parágrafo, como no corpus.
    /// </summary>
    public static byte[] WithTextAroundTopAnchoredImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(new Paragraph(
            new Run(new Text("Múltiplos")),
            new Run(AnchoredDrawing(part.GetIdOfPart(image))),
            new Run(new Text(" registros OK") { Space = SpaceProcessingModeValues.Preserve })));
    });

    /// <summary>
    /// Imagem que o LibreOffice ancora no lugar do próprio parágrafo: sem deslocamento,
    /// centralizada e com a largura da coluna, como num documento de capturas de tela.
    /// </summary>
    public static byte[] WithAnchoredImageInTheFlow() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(Paragraph("Antes da captura."));
        body.AppendChild(new Paragraph(new Run(AnchoredDrawing(part.GetIdOfPart(image)))));
        body.AppendChild(Paragraph("Depois da captura."));
    });

    /// <summary>Sete seções <c>continuous</c> de geometria idêntica, como o LibreOffice grava.</summary>
    public static byte[] WithIdenticalSections() => Build((body, _) =>
    {
        for (var i = 1; i <= 6; i++)
        {
            var paragraph = Paragraph($"Trecho {i}.");
            // A mesma geometria da seção final que `Build` acrescenta.
            paragraph.ParagraphProperties = new ParagraphProperties(
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.Continuous },
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U }));
            body.AppendChild(paragraph);
        }

        body.AppendChild(Paragraph("Fim."));
    });

    /// <summary>Duas seções que divergem: a primeira em paisagem.</summary>
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
    /// Três seções: a primeira declara cabeçalho e numeração em romanos; a do
    /// meio, contínua, herda o cabeçalho; a última (do corpo) começa em página
    /// ímpar, em paisagem, e reinicia a numeração em 1.
    /// </summary>
    public static byte[] WithThreeSections() => Build((body, part) =>
    {
        var header = part.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text("Cabeçalho da primeira"))));
        var headerId = part.GetIdOfPart(header);

        body.AppendChild(Paragraph("Primeira seção."));
        var first = new Paragraph();
        first.ParagraphProperties = new ParagraphProperties(
            new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId },
                new PageSize { Width = 11906U, Height = 16838U },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U },
                new PageNumberType { Format = NumberFormatValues.LowerRoman }));
        body.AppendChild(first);

        var second = Paragraph("Segunda seção, que termina aqui.");
        second.ParagraphProperties = new ParagraphProperties(
            new SectionProperties(
                new SectionType { Val = SectionMarkValues.Continuous },
                new PageSize { Width = 11906U, Height = 16838U },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U }));
        body.AppendChild(second);

        body.AppendChild(Paragraph("Terceira seção."));
    },
    (section, _) =>
    {
        section.PrependChild(new SectionType { Val = SectionMarkValues.OddPage });
        var size = section.GetFirstChild<PageSize>()!;
        size.Width = 16838U;
        size.Height = 11906U;
        size.Orient = PageOrientationValues.Landscape;
        section.AppendChild(new PageNumberType { Start = 1 });
    });

    /// <summary>
    /// Colunas: a primeira seção em duas colunas iguais com linha entre elas e
    /// uma quebra de coluna; a do corpo em três de larguras diferentes.
    /// </summary>
    public static byte[] WithColumns() => Build((body, _) =>
    {
        var broken = Paragraph("Fim da primeira coluna.");
        broken.AppendChild(new Run(new Break { Type = BreakValues.Column }));
        body.AppendChild(broken);
        body.AppendChild(Paragraph("Segunda coluna."));

        var mark = new Paragraph();
        mark.ParagraphProperties = new ParagraphProperties(
            new SectionProperties(
                new PageSize { Width = 11906U, Height = 16838U },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U },
                new Columns { ColumnCount = 2, Space = "567", Separator = true }));
        body.AppendChild(mark);
        body.AppendChild(Paragraph("Três colunas desiguais."));
    },
    (section, _) =>
    {
        section.PrependChild(new SectionType { Val = SectionMarkValues.Continuous });
        section.AppendChild(new Columns(
            new Column { Width = "2000", Space = "400" },
            new Column { Width = "3000", Space = "400" },
            new Column { Width = "3226" }) { ColumnCount = 3, EqualWidth = false });
    });

    /// <summary>
    /// Formatação nos estilos, como no corpus: <c>Faixa</c> é o <c>Heading1</c> dele
    /// (fundo vermelho, texto branco, Arial 10 pt, centralizado), e <c>Corpo</c> herda
    /// de <c>Base</c>.
    /// </summary>
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

    /// <summary>
    /// A medida da linha: o corpo sem <c>w:pStyle</c> no estilo <c>w:default="1"</c>, a
    /// fonte da marca de parágrafo (<c>w:pPr/w:rPr</c>), com que o Word mede a linha, e
    /// entrelinhas declaradas ou não.
    /// </summary>
    public static byte[] WithLineMetrics() => Build((body, part) =>
    {
        var styles = part.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Times New Roman" }))),

            new Style(
                new StyleName { Val = "Normal" },
                new StyleRunProperties(new FontSize { Val = "24" }))
            { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        // Sem `w:pStyle`: só o estilo padrão diz que isto é 12 pt.
        body.AppendChild(Paragraph("Herda o estilo padrão."));

        // A marca de parágrafo manda na altura da linha, e diverge do estilo.
        var marked = Paragraph("Verdana de dez pontos.");
        marked.ParagraphProperties = new ParagraphProperties(
            new ParagraphMarkRunProperties(
                new RunFonts { Ascii = "Verdana" },
                new FontSize { Val = "20" }));
        body.AppendChild(marked);

        // Entrelinha travada em 9 pt.
        var exact = Paragraph("Entrelinha travada.");
        exact.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "180", LineRule = LineSpacingRuleValues.Exact });
        body.AppendChild(exact);

        // Uma vez e meia, que é o outro valor que aparece na prática.
        var loose = Paragraph("Entrelinha de uma vez e meia.");
        loose.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto });
        body.AppendChild(loose);

        // 271/240 em Arial, o múltiplo e a fonte mais comuns do corpus.
        var multiple = Paragraph("Arial e um pouco mais de linha.");
        multiple.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "271", LineRule = LineSpacingRuleValues.Auto },
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Arial" }));
        body.AppendChild(multiple);

        // Fonte que o instalador não leva: a substituta depende da máquina.
        var unknown = Paragraph("Numa fonte que ninguém tem.");
        unknown.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "271", LineRule = LineSpacingRuleValues.Auto },
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Fonte Fantasma" }));
        body.AppendChild(unknown);
    });

    /// <summary>
    /// Estilo com entrelinha 276 e recuo, e parágrafos que redeclaram só
    /// <c>w:before</c> e <c>w:after</c>, como o documento de evidências do corpus.
    /// </summary>
    public static byte[] WithStyleSpacingAndDirectMargins() => Build((body, part) =>
    {
        var styles = part.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(
            new Style(
                new StyleName { Val = "Normal" },
                new StyleParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Line = "276",
                        LineRule = LineSpacingRuleValues.Auto,
                        Before = "0",
                        After = "140",
                    },
                    new Indentation { Left = "720", Right = "60", Hanging = "360" }),
                new StyleRunProperties(new RunFonts { Ascii = "Arial" }, new FontSize { Val = "20" }))
            { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });

        // Só o espaço; a entrelinha e o recuo continuam sendo os do estilo.
        var paragraph = Paragraph("Herda a entrelinha e o recuo do estilo.");
        paragraph.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "0" });
        body.AppendChild(paragraph);

        // Este troca o recuo da esquerda e mantém o resto.
        var moved = Paragraph("Troca só o recuo da esquerda.");
        moved.ParagraphProperties = new ParagraphProperties(new Indentation { Left = "1440" });
        body.AppendChild(moved);
    });

    /// <summary>
    /// Marca de seção no meio do texto, como o LibreOffice grava: o <c>w:sectPr</c> no
    /// <c>w:pPr</c> de um parágrafo vazio.
    /// </summary>
    public static byte[] WithSectionMarkInTheMiddle() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Antes da marca."));

        var mark = new Paragraph();
        mark.ParagraphProperties = new ParagraphProperties(
            new SectionProperties(
                new DocumentFormat.OpenXml.Wordprocessing.PageSize { Width = 11906U, Height = 16838U }));
        body.AppendChild(mark);

        body.AppendChild(Paragraph("Depois da marca."));
    });

    /// <summary>Um parágrafo que não se corta entre linhas, e um que desliga o controle de viúvas.</summary>
    public static byte[] WithKeepLines() => Build((body, _) =>
    {
        var kept = Paragraph("Linhas juntas.");
        kept.ParagraphProperties = new ParagraphProperties(new KeepLines());
        body.AppendChild(kept);

        var loose = Paragraph("Viúva permitida.");
        loose.ParagraphProperties = new ParagraphProperties(new WidowControl { Val = false });
        body.AppendChild(loose);

        body.AppendChild(Paragraph("Comum."));
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
    /// Alinhado à esquerda e centralizado por tabulação, como no corpus: <c>w:jc</c>
    /// <c>left</c> e uma parada centralizada no meio da coluna.
    /// </summary>
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

    /// <summary>Sobrescrito e subscrito: a fórmula e a nota de referência.</summary>
    public static byte[] WithVerticalAlignment() => Build((body, _) =>
    {
        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(new Text("H")));
        paragraph.AppendChild(new Run(
            new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }),
            new Text("2")));
        paragraph.AppendChild(new Run(new Text("O e m")));
        paragraph.AppendChild(new Run(
            new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new Text("2")));
        body.AppendChild(paragraph);
    });

    /// <summary>
    /// Duas tabelas: uma com `w:tblCellMar` próprio só em cima, outra calada — que
    /// herda do estilo padrão de tabela (60 twips embaixo) e do Word (108 dos lados).
    /// </summary>
    public static byte[] WithCellMargins() => Build((body, part) =>
    {
        var styles = part.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(new Style(
            new StyleName { Val = "Normal Table" },
            new StyleTableProperties(new TableCellMarginDefault(
                new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa })))
        {
            Type = StyleValues.Table,
            StyleId = "TableNormal",
            Default = true,
        });

        Table Make(string text, TableCellMarginDefault? margins) => new(
            new TableProperties(margins is null ? [] : [margins]),
            new TableGrid(new GridColumn { Width = "9000" }),
            new TableRow(new TableCell(Paragraph(text))));

        body.AppendChild(Make("Margem própria", new TableCellMarginDefault(
            new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa })));
        body.AppendChild(Paragraph("Entre as tabelas."));
        body.AppendChild(Make("Margem herdada", null));
    });

    public static byte[] WithTable() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Antes da tabela."));

        // O esquema exige `w:tblGrid` em toda tabela.
        var table = new Table(
            new TableProperties(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 })),
            new TableGrid(new GridColumn { Width = "4675" }, new GridColumn { Width = "4675" }));

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
        AddBulletNumbering(part);

        body.AppendChild(Paragraph("Introdução."));
        body.AppendChild(NumberedParagraph("Primeiro item", 1));
        body.AppendChild(NumberedParagraph("Segundo item", 1));
        body.AppendChild(Paragraph("Conclusão."));
    });

    /// <summary>
    /// Lista numerada de dois níveis: o segundo compõe o primeiro (<c>%1.%2)</c>), a
    /// lista continua depois de um parágrafo comum (o mesmo <c>numId</c>), e um segundo
    /// <c>w:num</c> com <c>w:startOverride</c> conta a partir de 10.
    /// </summary>
    public static byte[] WithMultilevelList() => Build((body, part) =>
    {
        Level Level(int index, NumberFormatValues format, string text, int left) => new(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = format },
            new LevelText { Val = text },
            new PreviousParagraphProperties(new Indentation { Left = $"{left}", Hanging = "360" }))
        {
            LevelIndex = index,
        };

        var numbering = part.AddNewPart<NumberingDefinitionsPart>();
        numbering.Numbering = new Numbering(
            new AbstractNum(
                Level(0, NumberFormatValues.Decimal, "%1.", 360),
                Level(1, NumberFormatValues.LowerLetter, "%1.%2)", 720)) { AbstractNumberId = 3 },
            new NumberingInstance(new AbstractNumId { Val = 3 }) { NumberID = 5 },
            new NumberingInstance(
                new AbstractNumId { Val = 3 },
                new LevelOverride(new StartOverrideNumberingValue { Val = 10 }) { LevelIndex = 0 }) { NumberID = 6 });

        body.AppendChild(Paragraph("Introdução."));
        body.AppendChild(NumberedParagraph("Um", 5));
        body.AppendChild(NumberedParagraph("Um-a", 5, 1));
        body.AppendChild(NumberedParagraph("Um-b", 5, 1));
        body.AppendChild(NumberedParagraph("Dois", 5));
        body.AppendChild(Paragraph("No meio."));
        body.AppendChild(NumberedParagraph("Três", 5));
        body.AppendChild(NumberedParagraph("Dez", 6));
        body.AppendChild(Paragraph("Conclusão."));
    });

    /// <summary>
    /// Lista com marcador dentro de uma célula, onde o editor vê parágrafos comuns: o
    /// <c>w:numPr</c> tem de sobreviver.
    /// </summary>
    public static byte[] WithListInsideTableCell() => Build((body, part) =>
    {
        AddBulletNumbering(part);

        var cell = new TableCell(NumberedParagraph("Primeiro da célula", 1));
        cell.AppendChild(NumberedParagraph("Segundo da célula", 1));

        var table = new Table(
            new TableProperties(new TableStyle { Val = "GradeComLista" }),
            new TableGrid(new GridColumn { Width = "9350" }));
        table.AppendChild(new TableRow(cell));

        body.AppendChild(table);
    });

    /// <summary>
    /// Recuo negativo, que o leitor não emite: o modelo volta dizendo zero sobre um
    /// recuo que existe.
    /// </summary>
    public static byte[] WithNegativeIndent() => Build((body, _) =>
    {
        var paragraph = Paragraph("Texto para fora da margem.");
        paragraph.ParagraphProperties = new ParagraphProperties(new Indentation { Left = "-284" });
        body.AppendChild(paragraph);
    });

    /// <summary>
    /// Tabela com um marcador entre duas linhas, filho legítimo de <c>w:tbl</c>.
    /// </summary>
    public static byte[] WithBookmarkBetweenRows() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(new TableStyle { Val = "GradeMarcada" }),
            new TableGrid(new GridColumn { Width = "9350" }));
        table.AppendChild(new TableRow(new TableCell(Paragraph("Linha de cima"))));
        table.AppendChild(new BookmarkStart { Id = "1", Name = "MeioDaTabela" });
        table.AppendChild(new TableRow(new TableCell(Paragraph("Linha de baixo"))));
        table.AppendChild(new BookmarkEnd { Id = "1" });

        body.AppendChild(table);
    });


    /// <summary>
    /// Parágrafo abraçado por um marcador, destino de referência cruzada, índice e
    /// link interno, que o editor não representa.
    /// </summary>
    public static byte[] WithBookmarkAroundParagraph() => Build((body, _) =>
    {
        var marked = new Paragraph();
        marked.AppendChild(new BookmarkStart { Id = "1", Name = "CapituloUm" });
        marked.AppendChild(new Run(new Text("Parágrafo marcado.") { Space = SpaceProcessingModeValues.Preserve }));
        marked.AppendChild(new BookmarkEnd { Id = "1" });

        body.AppendChild(marked);
        body.AppendChild(Paragraph("Parágrafo comum."));
    });

    /// <summary>
    /// Parágrafo com tudo o que o <c>w:pPr</c> diz: o que a tela mostra (estilo,
    /// espaçamento, entrelinha, fundo, fonte da marca) e o que não (borda, tabulação).
    /// </summary>
    public static byte[] WithFormattedParagraph() => Build((body, part) =>
    {
        var styles = part.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(new Style(
            new StyleName { Val = "Título 1" },
            new StyleParagraphProperties(new SpacingBetweenLines { Before = "240", After = "120" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Ttulo1",
        });

        var properties = new ParagraphProperties(
            new ParagraphStyleId { Val = "Ttulo1" },
            new KeepNext(),
            new ParagraphBorders(new BottomBorder { Val = BorderValues.Double, Size = 18 }),
            new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "C00000" },
            new Tabs(new TabStop { Val = TabStopValues.Center, Position = 4500 }),
            new SpacingBetweenLines
            {
                Before = "360",
                After = "180",
                Line = "271",
                LineRule = LineSpacingRuleValues.Auto,
            },
            new Indentation { Left = "720", FirstLine = "360" },
            new Justification { Val = JustificationValues.Center },
            new ParagraphMarkRunProperties(
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize { Val = "20" }));

        var paragraph = new Paragraph(properties);
        paragraph.AppendChild(new Run(new Text("Título formatado.") { Space = SpaceProcessingModeValues.Preserve }));

        body.AppendChild(paragraph);
        body.AppendChild(Paragraph("Parágrafo comum."));
    });

    /// <summary>
    /// Tabela com o que o modelo não representa: estilo, largura, grade, cabeçalho que
    /// se repete, sombreamento e mesclagem vertical.
    /// </summary>
    public static byte[] WithStyledTable() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(
                new TableStyle { Val = "GradeMedia3" },
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Double, Size = 18 },
                    new BottomBorder { Val = BorderValues.Double, Size = 18 })),
            new TableGrid(
                new GridColumn { Width = "4000" },
                new GridColumn { Width = "5000" }));

        var header = new TableRow(new TableRowProperties(new TableHeader()));
        header.AppendChild(new TableCell(
            new TableCellProperties(
                new TableCellWidth { Width = "4000", Type = TableWidthUnitValues.Dxa },
                new VerticalMerge { Val = MergedCellValues.Restart },
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9D9D9" }),
            Paragraph("Cabeçalho A")));
        header.AppendChild(new TableCell(Paragraph("Cabeçalho B")));
        table.AppendChild(header);

        var row = new TableRow();
        row.AppendChild(new TableCell(
            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Continue }),
            Paragraph("Continuação")));
        row.AppendChild(new TableCell(Paragraph("Dado B")));
        table.AppendChild(row);

        body.AppendChild(table);
    });

    /// <summary>
    /// Trama de sombreamento (<c>pct25</c>) e borda <c>thickThinSmallGap</c>, que a tela
    /// aproxima por cor lisa e linha simples.
    /// </summary>
    public static byte[] WithPatternedCell() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "4500" }, new GridColumn { Width = "4500" }));

        var row = new TableRow();
        row.AppendChild(new TableCell(
            new TableCellProperties(
                new TableCellBorders(new TopBorder { Val = BorderValues.ThickThinSmallGap, Size = 24 }),
                new Shading { Val = ShadingPatternValues.Percent25, Color = "FF0000", Fill = "FFFF00" }),
            Paragraph("Com trama")));
        row.AppendChild(new TableCell(Paragraph("Sem trama")));
        table.AppendChild(row);

        body.AppendChild(table);
    });

    /// <summary>
    /// Célula com as bordas que o modelo não representa: diagonal, borda
    /// interna, `w:space` e cor de tema num lado, e o `w:nil` de baixo que o Word
    /// grava o tempo todo.
    /// </summary>
    public static byte[] WithRichCellBorders() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "4500" }, new GridColumn { Width = "4500" }));

        var row = new TableRow();
        row.AppendChild(new TableCell(
            new TableCellProperties(
                new TableCellBorders(
                    new TopBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4,
                        Space = 0,
                        Color = "4472C4",
                        ThemeColor = ThemeColorValues.Accent1,
                    },
                    new BottomBorder { Val = BorderValues.Nil },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new TopLeftToBottomRightCellBorder { Val = BorderValues.Single, Size = 4 })),
            Paragraph("Com diagonal")));
        row.AppendChild(new TableCell(Paragraph("Sem borda")));
        table.AppendChild(row);

        body.AppendChild(table);
    });

    /// <summary>
    /// Tabela cuja primeira linha não cobre a grade: começa uma coluna adiante
    /// (`w:gridBefore`), como em formulário e em documento convertido de PDF.
    /// </summary>
    public static byte[] WithGridBefore() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(
                new GridColumn { Width = "2000" },
                new GridColumn { Width = "3000" },
                new GridColumn { Width = "4000" }));

        var first = new TableRow(new TableRowProperties(new GridBefore { Val = 1 }));
        first.AppendChild(new TableCell(Paragraph("Recuada B")));
        first.AppendChild(new TableCell(Paragraph("Recuada C")));
        table.AppendChild(first);

        var second = new TableRow();
        second.AppendChild(new TableCell(Paragraph("Cheia A")));
        second.AppendChild(new TableCell(Paragraph("Cheia B")));
        second.AppendChild(new TableCell(Paragraph("Cheia C")));
        table.AppendChild(second);

        body.AppendChild(table);
    });

    /// <summary>
    /// Tabela de três colunas de larguras diferentes, com medidas que não são
    /// múltiplo de 15 twips — o pixel do CSS. 2000 twips são 133,33 px.
    /// </summary>
    public static byte[] WithThreeColumns() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(
                new GridColumn { Width = "2000" },
                new GridColumn { Width = "3000" },
                new GridColumn { Width = "4000" }));

        foreach (var prefix in new[] { "Um", "Dois" })
        {
            var row = new TableRow();
            foreach (var column in new[] { "A", "B", "C" })
            {
                row.AppendChild(new TableCell(Paragraph($"{prefix} {column}")));
            }

            table.AppendChild(row);
        }

        body.AppendChild(table);
    });

    /// <summary>
    /// Tabela cuja primeira linha foi inserida com o controle de alterações
    /// ligado: o `w:ins` mora no fim do `w:trPr`.
    /// </summary>
    public static byte[] WithInsertedRow() => Build((body, _) =>
    {
        var table = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "4500" }, new GridColumn { Width = "4500" }));

        var first = new TableRow(new TableRowProperties(
            new TableRowHeight { Val = 400U },
            new Inserted { Id = "1", Author = "Revisora", Date = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc) }));
        first.AppendChild(new TableCell(Paragraph("Título A")));
        first.AppendChild(new TableCell(Paragraph("Título B")));
        table.AppendChild(first);

        var second = new TableRow();
        second.AppendChild(new TableCell(Paragraph("Dado A")));
        second.AppendChild(new TableCell(Paragraph("Dado B")));
        table.AppendChild(second);

        body.AppendChild(table);
    });

    /// <summary>
    /// Tabela com uma célula inserida pelo controle de alterações (`w:cellIns`):
    /// revisão de estrutura, que o editor não representa.
    /// </summary>
    public static byte[] WithInsertedCell() => BuildFromXml(
        """
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:cellIns w:id="1" w:author="Revisora"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl><w:p/>
        """,
        string.Empty);

    /// <summary>Tabela dentro de uma célula de outra tabela.</summary>
    public static byte[] WithNestedTable() => Build((body, _) =>
    {
        var inner = new Table(
            new TableProperties(new TableStyle { Val = "GradeInterna" }),
            new TableGrid(new GridColumn { Width = "4000" }));
        inner.AppendChild(new TableRow(new TableCell(Paragraph("Dentro da tabela de dentro"))));

        var outer = new Table(
            new TableProperties(new TableStyle { Val = "GradeExterna" }),
            new TableGrid(new GridColumn { Width = "9350" }));
        var cell = new TableCell(Paragraph("Antes da aninhada"));
        cell.AppendChild(inner);
        // `w:tc` não pode terminar em tabela: o Word exige um parágrafo depois.
        cell.AppendChild(Paragraph("Depois da aninhada"));
        outer.AppendChild(new TableRow(cell));

        body.AppendChild(outer);
    });

    /// <summary>Papel A5, que o modelo do editor não nomeia.</summary>
    public static byte[] WithCustomPaper() => Build(
        (body, _) => body.AppendChild(Paragraph("Meia folha.")),
        (section, _) =>
        {
            var size = section.GetFirstChild<PageSize>()!;
            size.Width = 8391U;
            size.Height = 11907U;
        });

    /// <summary>Um PNG de 4 × 4, para provar que a medida sai do cabeçalho dele.</summary>
    public static byte[] SquarePng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAADklEQVR4nGNwQAIMxHEAOEMMAfoZu1cAAAAASUVORK5CYII=");



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

    /// <summary>
    /// Numeração com um marcador: o quadrado da Wingdings, na área de uso privado.
    /// </summary>
    private static void AddBulletNumbering(MainDocumentPart part)
    {
        var level = new Level(
            new NumberingFormat { Val = NumberFormatValues.Bullet },
            new LevelText { Val = "\uF0A7" },
            new PreviousParagraphProperties(new Indentation { Left = "720", Hanging = "360" }))
        {
            LevelIndex = 0,
        };

        var numbering = part.AddNewPart<NumberingDefinitionsPart>();
        numbering.Numbering = new Numbering(
            new AbstractNum(level) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
    }

    /// <summary>No arquivo, item de lista é parágrafo que aponta uma numeração.</summary>
    private static Paragraph NumberedParagraph(string text, int numId, int level = 0)
    {
        var paragraph = Paragraph(text);
        paragraph.ParagraphProperties = new ParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = numId }));
        return paragraph;
    }

    /// <summary>
    /// Um parágrafo, a seção decorada por quem pede e um <c>settings.xml</c> mínimo,
    /// para conferir que a gravação não o toca.
    /// </summary>
    public static byte[] WithSection(Action<SectionProperties> decorate) => Build(
        (body, part) =>
        {
            body.AppendChild(Paragraph("Corpo do documento."));
            var settings = part.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new Settings(new DefaultTabStop { Val = 708 });
            settings.Settings.Save();
        },
        (section, _) => decorate(section));

    /// <summary>Um rodapé padrão com o parágrafo dado — para campos e textos de faixa.</summary>
    public static byte[] WithFooter(Paragraph paragraph) => Build(
        (body, _) => body.AppendChild(Paragraph("Corpo do documento.")),
        (section, part) =>
        {
            var footer = part.AddNewPart<FooterPart>("rIdRodape");
            footer.Footer = new Footer(paragraph);
            footer.Footer.Save();
            section.AppendChild(new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdRodape" });
        });

    /// <summary>
    /// Parágrafos numerados por um `numId` que o `numbering.xml` não define — o
    /// Word os mostra sem marca nenhuma.
    /// </summary>
    public static byte[] WithDanglingNumbering() => Build((body, part) =>
    {
        AddBulletNumbering(part);
        body.AppendChild(Paragraph("Introdução."));
        body.AppendChild(NumberedParagraph("Sem definição", 99));
        body.AppendChild(NumberedParagraph("Também sem", 99));
    });

    /// <summary>
    /// Lista cujo nível 0 tem o que a definição do editor não leva: número à
    /// direita, em vermelho, e formato `ordinal`.
    /// </summary>
    public static byte[] WithRichNumbering() => Build((body, part) =>
    {
        var numbering = part.AddNewPart<NumberingDefinitionsPart>();
        numbering.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Ordinal },
                    new LevelText { Val = "%1" },
                    new LevelJustification { Val = LevelJustificationValues.Right },
                    new PreviousParagraphProperties(new Indentation { Left = "720", Hanging = "360" }),
                    new NumberingSymbolRunProperties(new Color { Val = "FF0000" }))
                { LevelIndex = 0 },
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                    new LevelText { Val = "%2." },
                    new PreviousParagraphProperties(new Indentation { Left = "1440", Hanging = "360" }))
                { LevelIndex = 1 }) { AbstractNumberId = 4 },
            new NumberingInstance(new AbstractNumId { Val = 4 }) { NumberID = 1 });

        body.AppendChild(NumberedParagraph("Primeiro", 1));
        body.AppendChild(NumberedParagraph("Segundo", 1));
    });

    /// <summary>
    /// As referências como o Word as grava, em XML cru: sumário num controle de
    /// conteúdo, títulos com <c>_Toc…</c>, marcador que termina entre dois parágrafos,
    /// <c>SEQ</c>, <c>REF</c>, <c>PAGEREF</c> e link interno, com campo partido em cinco
    /// runs, dentro de link e entre parágrafos.
    /// </summary>
    public static byte[] WithReferences() => BuildFromXml(ReferencesBody, ReferencesStyles);

    internal const string ReferencesBody = """
        <w:sdt><w:sdtPr><w:id w:val="-1"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr/><w:sdtContent>
        <w:p><w:pPr><w:pStyle w:val="CabealhodoSumrio"/></w:pPr><w:r><w:t>Sumário</w:t></w:r></w:p>
        <w:p><w:pPr><w:pStyle w:val="Sumrio1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc100" w:history="1"><w:r><w:t>Introdução</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc100 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>
        <w:p><w:pPr><w:pStyle w:val="Sumrio2"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs></w:pPr><w:hyperlink w:anchor="_Toc101" w:history="1"><w:r><w:t>Escopo</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc101 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>
        <w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc100"/><w:r><w:t>Introdução</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
        <w:p><w:bookmarkStart w:id="1" w:name="Resumo"/><w:r><w:t xml:space="preserve">O resumo começa aqui </w:t></w:r></w:p>
        <w:p><w:r><w:t>e termina aqui.</w:t></w:r></w:p>
        <w:bookmarkEnd w:id="1"/><w:p><w:pPr><w:pStyle w:val="Ttulo2"/></w:pPr><w:bookmarkStart w:id="2" w:name="_Toc101"/><w:r><w:t>Escopo</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>
        <w:p><w:pPr><w:pStyle w:val="Legenda"/></w:pPr><w:bookmarkStart w:id="3" w:name="_Ref200"/><w:r><w:t xml:space="preserve">Figura </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figura \* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="3"/><w:r><w:t xml:space="preserve"> — Arquitetura</w:t></w:r></w:p>
        <w:p><w:r><w:t xml:space="preserve">Como mostra a </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref200 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Figura 1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, na página </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Ref200 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">. Veja o </w:t></w:r><w:hyperlink w:anchor="Resumo" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>resumo</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>
        """;

    internal const string ReferencesStyles = """
        <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
        <w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Ttulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="CabealhodoSumrio"><w:name w:val="TOC Heading"/><w:basedOn w:val="Ttulo1"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Sumrio1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:after="100"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Sumrio2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr></w:style>
        <w:style w:type="paragraph" w:styleId="Legenda"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="200"/></w:pPr><w:rPr><w:i/><w:sz w:val="18"/></w:rPr></w:style>
        <w:style w:type="character" w:default="1" w:styleId="Fontepargpadro"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
        <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="Fontepargpadro"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
        """;

    /// <summary>Um pacote com o corpo e os estilos dados em XML.</summary>
    internal static byte[] BuildFromXml(string body, string styles)
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var part = document.AddMainDocumentPart();
            var stylepart = part.AddNewPart<StyleDefinitionsPart>();
            stylepart.Styles = new Styles($"<w:styles {ns}>{styles}</w:styles>");
            stylepart.Styles.Save();
            part.Document = new Document(
                $"<w:document {ns}><w:body>{body}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
                "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>");
            part.Document.Save();
        }

        return buffer.ToArray();
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

            // O `w:sectPr` abre pelas referências de faixa: senão o fixture sai fora do esquema.
            var references = section.ChildElements
                .Where(child => child is HeaderReference or FooterReference)
                .ToList();
            for (var index = 0; index < references.Count; index++)
            {
                references[index].Remove();
                section.InsertAt(references[index], index);
            }
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
    private static OpenXmlElement TextBoxShape(string text, string decoration = "")
    {
        var xml = $"""
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
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
                            {decoration}
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

        // Pelo XML completo: `InnerXml` perderia o elemento que envolve os ramos.
        return new AlternateContent(xml);
    }

    /// <summary>Imagem **no fluxo** (`wp:inline`): ocupa lugar na linha.</summary>
    /// <param name="docPrId">
    /// O `wp:docPr/@id` do desenho que já está no arquivo. O fixture o escolhe
    /// porque é com ele que o id de uma imagem nova não pode colidir.
    /// </param>
    public static byte[] WithInlineImage(uint docPrId = 1) => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(new Paragraph(new Run(InlineDrawing(part.GetIdOfPart(image), docPrId))));
    });

    /// <summary>Imagem no fluxo **com texto alternativo** (`wp:docPr/@descr`).</summary>
    public static byte[] WithDescribedImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        var drawing = InlineDrawing(part.GetIdOfPart(image));
        drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
            .First()
            .Description = "Organograma da diretoria";

        body.AppendChild(new Paragraph(new Run(drawing)));
    });

    private static OpenXmlElement InlineDrawing(string relationshipId, uint docPrId = 1)
    {
        var xml = $"""
            <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       distT="0" distB="0" distL="0" distR="0">
              <wp:extent cx="5274000" cy="2637000"/>
              <wp:docPr id="{docPrId}" name="Imagem {docPrId}"/>
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

    /// <param name="placed">
    /// Fora do fluxo, como a marca da capa. Com <c>false</c>, a imagem que o
    /// LibreOffice ancora onde o fluxo já a poria.
    /// </param>
    private static OpenXmlElement AnchoredDrawing(
        string relationshipId,
        long cx = 5274000,
        long cy = 2637000,
        int rotation = 0,
        bool placed = false)
    {
        var position = placed
            ? """
                  <wp:positionH relativeFrom="column"><wp:posOffset>-4559425</wp:posOffset></wp:positionH>
                  <wp:positionV relativeFrom="paragraph"><wp:posOffset>2095009</wp:posOffset></wp:positionV>
              """
            : """
                  <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
                  <wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV>
              """;
        var wrap = placed ? "<wp:wrapNone/>" : """<wp:wrapSquare wrapText="bothSides"/>""";

        const string Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        var xml = $"""
            <wp:anchor xmlns:wp="{Wp}"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2"
                       behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">
              <wp:simplePos x="0" y="0"/>
            {position}
              <wp:extent cx="{cx}" cy="{cy}"/>
              {wrap}
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

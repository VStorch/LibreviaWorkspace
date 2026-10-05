using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Tests;

/// <summary>
/// Test documents built in code, readable in review. They reproduce the structures of the corpus,
/// which is not in the repository, with invented content.
/// </summary>
public static class Fixtures
{
    /// <summary>A simple document with three paragraphs.</summary>
    public static byte[] Simple() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Primeiro parágrafo.", style: "Heading1"));
        body.AppendChild(Paragraph("Segundo parágrafo, com texto comum."));
        body.AppendChild(Paragraph("Terceiro parágrafo."));
    });

    /// <summary>A comment anchored on the second paragraph.</summary>
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
    /// A thread, a resolved comment and a point comment, as Word writes them: the reply embraces
    /// the parent's range, and <c>commentsExtended.xml</c> links both through <c>w14:paraId</c>.
    /// The point comment only has the reference.
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

    /// <summary>A document with tracked changes in the second paragraph.</summary>
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

    /// <summary>An anchored, centered image, as LibreOffice writes it.</summary>
    public static byte[] WithAnchoredImage() => Build((body, part) =>
    {
        var image = part.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(TinyPng()))
        {
            image.FeedData(stream);
        }

        body.AppendChild(Paragraph("Antes da imagem."));
        // With a real position: outside the column, and the text runs underneath.
        body.AppendChild(new Paragraph(new Run(
            AnchoredDrawing(part.GetIdOfPart(image), placed: true))));
        body.AppendChild(Paragraph("Depois da imagem."));
    });

    /// <summary>
    /// A logo anchored to the right in the header. The position lives in the anchor, and the
    /// drawing's <c>a:off</c> is zero, as in every single-piece drawing.
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
    /// A page break at the end of a <c>w:r</c>, inside the paragraph, distinct from one that takes
    /// a paragraph of its own.
    /// </summary>
    public static byte[] WithBreakInsideParagraph() => Build((body, _) =>
    {
        var comQuebra = new Paragraph();
        comQuebra.AppendChild(new Run(new Text("Fim da primeira página.")));
        comQuebra.AppendChild(new Run(new Break { Type = BreakValues.Page }));
        body.AppendChild(comQuebra);
        body.AppendChild(Paragraph("Começo da segunda."));
    });

    /// <summary>A footer of three centered paragraphs, like the manual template's.</summary>
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
    /// A footer with text, a tab and the <c>PAGE</c> field: only the text has a <c>w:t</c> to write
    /// into.
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

    /// <summary>A lone page break in a paragraph.</summary>
    public static byte[] WithLonePageBreak() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Primeira página."));
        body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        body.AppendChild(Paragraph("Segunda página."));
    });

    /// <summary>
    /// Three headers, with <c>first</c> before <c>default</c> in the XML: the type counts, not the
    /// write order.
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
    /// A grid header, like the corpus corporate one: four columns and three rows, the logo merged
    /// vertically in the first and two columns joined by <c>w:gridSpan</c> on the right.
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

            // Without the bottom border, two file rows become a single frame.
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
    /// A font the machine may lack, with its kind declared in <c>word/fontTable.xml</c>.
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

        // A font the table does not declare: nothing to invent, it goes out as is.
        var outro = new Paragraph();
        outro.AppendChild(new Run(
            new RunProperties(new RunFonts { Ascii = "Fonte Fantasma", HighAnsi = "Fonte Fantasma" }),
            new Text("Sem tipo declarado") { Space = SpaceProcessingModeValues.Preserve }));
        body.AppendChild(outro);
    });

    /// <summary>
    /// A header that is a shape group, logo and title box: the anchor gives the group's position
    /// and size, and <c>a:chOff</c>/<c>a:chExt</c> the inner ruler.
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

    /// <summary>A new header with a line of text; returns the `r:id`.</summary>
    private static string Header(MainDocumentPart part, string text)
    {
        var header = part.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text(text))));
        header.Header.Save();
        return part.GetIdOfPart(header);
    }

    /// <summary>
    /// Title and subtitle in two boxes anchored on the same paragraph, like the manual template
    /// cover; each in <c>mc:Choice</c> and in <c>mc:Fallback</c>.
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
    /// Three boxes: undecorated (like the corpus ones, <c>a:noFill</c> and zero line), one that can
    /// be drawn and one that cannot.
    /// </summary>
    public static byte[] WithDecoratedTextBoxes() => Build((body, _) =>
    {
        // No frame, like the corpus header boxes.
        body.AppendChild(new Paragraph(new Run(TextBoxShape(
            "Sem moldura",
            """<a:noFill/><a:ln w="0"><a:noFill/></a:ln>"""))));

        // Solid fill and stroke: what CSS draws.
        body.AppendChild(new Paragraph(new Run(TextBoxShape(
            "Com moldura",
            """
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            <a:ln w="12700"><a:solidFill><a:srgbClr val="1F5FA9"/></a:solidFill></a:ln>
            """))));

        body.AppendChild(Paragraph("Corpo do documento."));
    });

    /// <summary>A box with a gradient fill, which we cannot draw.</summary>
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
    /// The cover's vertical mark: an anchored image rotated a quarter turn. <c>wp:extent</c>
    /// measures the image lying down, 28.58 × 8.01 cm.
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
    /// How the manual template ends its cover: a paragraph with only the positioned mark and the
    /// <c>w:br</c> to the next sheet.
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
    /// Text, a capture anchored to the top of the paragraph and more text, in the same paragraph,
    /// as in the corpus.
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
    /// An image LibreOffice anchors in its own paragraph's place: no offset, centered and
    /// column-wide, as in a screenshot document.
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

    /// <summary>
    /// Seven <c>continuous</c> sections with identical geometry, as LibreOffice writes them.
    /// </summary>
    public static byte[] WithIdenticalSections() => Build((body, _) =>
    {
        for (var i = 1; i <= 6; i++)
        {
            var paragraph = Paragraph($"Trecho {i}.");
            // The same geometry as the final section `Build` adds.
            paragraph.ParagraphProperties = new ParagraphProperties(
                new SectionProperties(
                    new SectionType { Val = SectionMarkValues.Continuous },
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U }));
            body.AppendChild(paragraph);
        }

        body.AppendChild(Paragraph("Fim."));
    });

    /// <summary>Two diverging sections: the first in landscape.</summary>
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
    /// Three sections: the first declares a header and Roman numbering; the middle one, continuous,
    /// inherits the header; the last (the body's) starts on an odd page, in landscape, and restarts
    /// numbering at 1.
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
    /// Columns: the first section in two equal columns with a line between them and a column break;
    /// the body's in three of different widths.
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
    /// Formatting in styles, as in the corpus: <c>Faixa</c> is its <c>Heading1</c> (red background,
    /// white text, Arial 10 pt, centered), and <c>Corpo</c> inherits from <c>Base</c>.
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

        // Direct formatting must beat the style.
        var overridden = Paragraph("Alinhado à direita.", style: "Corpo");
        overridden.ParagraphProperties!.AppendChild(new Justification { Val = JustificationValues.Right });
        body.AppendChild(overridden);

        // `w:rPr` inside `w:pPr` formats the paragraph mark, not the runs.
        var marked = Paragraph("Não deve ficar em negrito.", style: "Corpo");
        marked.ParagraphProperties!.AppendChild(new ParagraphMarkRunProperties(new Bold()));
        body.AppendChild(marked);
    });

    /// <summary>
    /// Line measurement: a body without <c>w:pStyle</c> in the <c>w:default="1"</c> style, the
    /// paragraph mark font (<c>w:pPr/w:rPr</c>) Word measures the line with, and declared or
    /// undeclared line spacing.
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

        // Without `w:pStyle`: only the default style says this is 12 pt.
        body.AppendChild(Paragraph("Herda o estilo padrão."));

        // The paragraph mark rules line height, and differs from the style.
        var marked = Paragraph("Verdana de dez pontos.");
        marked.ParagraphProperties = new ParagraphProperties(
            new ParagraphMarkRunProperties(
                new RunFonts { Ascii = "Verdana" },
                new FontSize { Val = "20" }));
        body.AppendChild(marked);

        // Line spacing fixed at 9 pt.
        var exact = Paragraph("Entrelinha travada.");
        exact.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "180", LineRule = LineSpacingRuleValues.Exact });
        body.AppendChild(exact);

        // One and a half, the other value seen in practice.
        var loose = Paragraph("Entrelinha de uma vez e meia.");
        loose.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto });
        body.AppendChild(loose);

        // 271/240 in Arial, the corpus's most common multiple and font.
        var multiple = Paragraph("Arial e um pouco mais de linha.");
        multiple.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "271", LineRule = LineSpacingRuleValues.Auto },
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Arial" }));
        body.AppendChild(multiple);

        // A font the installer does not ship: the substitute depends on the machine.
        var unknown = Paragraph("Numa fonte que ninguém tem.");
        unknown.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "271", LineRule = LineSpacingRuleValues.Auto },
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Fonte Fantasma" }));
        body.AppendChild(unknown);
    });

    /// <summary>
    /// A style with 276 line spacing and an indent, and paragraphs redeclaring only <c>w:before</c>
    /// and <c>w:after</c>, like the corpus evidence document.
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

        // Only the space; line spacing and indent stay the style's.
        var paragraph = Paragraph("Herda a entrelinha e o recuo do estilo.");
        paragraph.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "0" });
        body.AppendChild(paragraph);

        // This one changes the left indent and keeps the rest.
        var moved = Paragraph("Troca só o recuo da esquerda.");
        moved.ParagraphProperties = new ParagraphProperties(new Indentation { Left = "1440" });
        body.AppendChild(moved);
    });

    /// <summary>
    /// A section mark in the middle of the text, as LibreOffice writes it: <c>w:sectPr</c> in the
    /// <c>w:pPr</c> of an empty paragraph.
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

    /// <summary>
    /// A paragraph that does not break between lines, and one that turns widow control off.
    /// </summary>
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

    /// <summary>A paragraph that asks to stay with the next, another that does not.</summary>
    public static byte[] WithKeepNext() => Build((body, _) =>
    {
        var kept = Paragraph("Rótulo da imagem:");
        kept.ParagraphProperties = new ParagraphProperties(new KeepNext());
        body.AppendChild(kept);

        body.AppendChild(Paragraph("Solto no meio do texto."));
    });

    /// <summary>
    /// Left-aligned and centered by a tab, as in the corpus: <c>w:jc</c> <c>left</c> and a centered
    /// stop in the middle of the column.
    /// </summary>
    public static byte[] WithTabCentering() => Build((body, _) =>
    {
        var centered = new Paragraph();
        centered.ParagraphProperties = new ParagraphProperties(
            new Tabs(new TabStop { Val = TabStopValues.Center, Position = 4153 }),
            new Justification { Val = JustificationValues.Left });
        centered.AppendChild(new Run(new TabChar(), new TabChar(), new Text("Centralizado por tabulação")));
        body.AppendChild(centered);

        // A tab mid-line is not paragraph positioning.
        var inline = new Paragraph();
        inline.ParagraphProperties = new ParagraphProperties(
            new Tabs(new TabStop { Val = TabStopValues.Center, Position = 4153 }),
            new Justification { Val = JustificationValues.Left });
        inline.AppendChild(new Run(new Text("Esquerda"), new TabChar(), new Text("Meio")));
        body.AppendChild(inline);
    });

    /// <summary>A style inheriting from itself: it must not hang the reader.</summary>
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

    /// <summary>Small caps and all caps: 45 occurrences in the corpus.</summary>
    public static byte[] WithSmallCaps() => Build((body, _) =>
    {
        var paragraph = new Paragraph();
        paragraph.AppendChild(new Run(new RunProperties(new SmallCaps()), new Text("versalete")));
        paragraph.AppendChild(new Run(new RunProperties(new Caps()), new Text("maiúsculas")));
        body.AppendChild(paragraph);
    });

    /// <summary>Superscript and subscript: the formula and the reference note.</summary>
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
    /// Two tables: one with its own `w:tblCellMar` only on top, another silent one, which inherits
    /// from the default table style (60 twips below) and from Word (108 on the sides).
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

        // The schema requires `w:tblGrid` in every table.
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

    /// <summary>A bullet list, with its numbering really declared.</summary>
    public static byte[] WithBulletList() => Build((body, part) =>
    {
        AddBulletNumbering(part);

        body.AppendChild(Paragraph("Introdução."));
        body.AppendChild(NumberedParagraph("Primeiro item", 1));
        body.AppendChild(NumberedParagraph("Segundo item", 1));
        body.AppendChild(Paragraph("Conclusão."));
    });

    /// <summary>
    /// A two-level numbered list: the second composes the first (<c>%1.%2)</c>), the list continues
    /// after a regular paragraph (the same <c>numId</c>), and a second <c>w:num</c> with
    /// <c>w:startOverride</c> counts from 10.
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
    /// A bullet list inside a cell, where the editor sees regular paragraphs: the <c>w:numPr</c>
    /// must survive.
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
    /// A negative indent, which the reader does not emit: the model comes back saying zero over an
    /// indent that exists.
    /// </summary>
    public static byte[] WithNegativeIndent() => Build((body, _) =>
    {
        var paragraph = Paragraph("Texto para fora da margem.");
        paragraph.ParagraphProperties = new ParagraphProperties(new Indentation { Left = "-284" });
        body.AppendChild(paragraph);
    });

    /// <summary>
    /// A table with a bookmark between two rows, a legitimate child of <c>w:tbl</c>.
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
    /// A paragraph embraced by a bookmark, the target of cross-references, indexes and internal
    /// links, which the editor does not represent.
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
    /// A paragraph with everything <c>w:pPr</c> says: what the screen shows (style, spacing, line
    /// spacing, background, mark font) and what it does not (border, tab).
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
    /// A table with what the model does not represent: style, width, grid, repeated header, shading
    /// and vertical merge.
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
    /// A shading pattern (<c>pct25</c>) and a <c>thickThinSmallGap</c> border, which the screen
    /// approximates with a flat color and a single line.
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
    /// A cell with borders the model does not represent: diagonal, inner border, `w:space` and a
    /// theme color on one side, and the bottom `w:nil` Word writes all the time.
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
    /// A table whose first row does not cover the grid: it starts one column ahead
    /// (`w:gridBefore`), as in forms and documents converted from PDF.
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
    /// A three-column table of different widths, with measures that are not multiples of 15 twips,
    /// the CSS pixel. 2000 twips is 133.33 px.
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
    /// A table whose first row was inserted with tracking on: the `w:ins` lives at the end of
    /// `w:trPr`.
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
    /// A table with a cell inserted by tracking (`w:cellIns`): a structure revision the editor does
    /// not represent.
    /// </summary>
    public static byte[] WithInsertedCell() => BuildFromXml(
        """
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:cellIns w:id="1" w:author="Revisora"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl><w:p/>
        """,
        string.Empty);

    /// <summary>A table inside a cell of another table.</summary>
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
        // A `w:tc` cannot end in a table: Word requires a paragraph after it.
        cell.AppendChild(Paragraph("Depois da aninhada"));
        outer.AppendChild(new TableRow(cell));

        body.AppendChild(outer);
    });

    /// <summary>A5 paper, which the editor model does not name.</summary>
    public static byte[] WithCustomPaper() => Build(
        (body, _) => body.AppendChild(Paragraph("Meia folha.")),
        (section, _) =>
        {
            var size = section.GetFirstChild<PageSize>()!;
            size.Width = 8391U;
            size.Height = 11907U;
        });

    /// <summary>A 4 × 4 PNG, to prove the measure comes from its header.</summary>
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

    /// <summary>Bullet numbering: the Wingdings square, in the private use area.</summary>
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

    /// <summary>In the file, a list item is a paragraph pointing to a numbering.</summary>
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
    /// A paragraph, the section decorated by the caller and a minimal <c>settings.xml</c>, to check
    /// that saving does not touch it.
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

    /// <summary>A default footer with the given paragraph, for band fields and texts.</summary>
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
    /// Paragraphs numbered by a `numId` that `numbering.xml` does not define: Word shows them
    /// without any marker.
    /// </summary>
    public static byte[] WithDanglingNumbering() => Build((body, part) =>
    {
        AddBulletNumbering(part);
        body.AppendChild(Paragraph("Introdução."));
        body.AppendChild(NumberedParagraph("Sem definição", 99));
        body.AppendChild(NumberedParagraph("Também sem", 99));
    });

    /// <summary>
    /// A list whose level 0 has what the editor definition does not carry: a right-aligned number,
    /// in red, with the `ordinal` format.
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
    /// References as Word writes them, in raw XML: a table of contents in a content control,
    /// headings with <c>_Toc…</c>, a bookmark ending between two paragraphs, <c>SEQ</c>,
    /// <c>REF</c>, <c>PAGEREF</c> and an internal link, with a field split across five runs, inside
    /// a link and between paragraphs.
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

    /// <summary>A package with the given body and styles in XML.</summary>
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
    /// Receives the already built `w:sectPr`, so the fixture can add band references or section
    /// switches.
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

            // `w:sectPr` opens with band references: otherwise the fixture falls outside the
            // schema.
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
    /// A text box the way Word writes it: the same text in DrawingML and in the VML fallback.
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

        // Through the full XML: `InnerXml` would lose the element wrapping the branches.
        return new AlternateContent(xml);
    }

    /// <summary>An image **in the flow** (`wp:inline`): it takes room in the line.</summary>
    /// <param name="docPrId">
    /// The `wp:docPr/@id` of the drawing already in the file. The fixture picks it because a new
    /// image's id must not collide with it.
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

    /// <summary>An image in the flow **with alt text** (`wp:docPr/@descr`).</summary>
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
    /// Out of the flow, like the cover mark. With <c>false</c>, the image LibreOffice anchors where
    /// the flow would already place it.
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

    /// <summary>A transparent 1×1 PNG, the smallest valid one.</summary>
    private static byte[] TinyPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
}

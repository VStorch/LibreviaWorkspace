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
        // Com posição de verdade: fora da coluna, e o texto passa por baixo.
        body.AppendChild(new Paragraph(new Run(
            AnchoredDrawing(part.GetIdOfPart(image), placed: true))));
        body.AppendChild(Paragraph("Depois da imagem."));
    });

    /// <summary>
    /// Cabeçalho com um logotipo ancorado à direita da coluna de texto.
    /// </summary>
    /// <remarks>
    /// A posição mora na âncora, e o `a:off` de dentro do desenho é zero —
    /// que é o caso de todo desenho de peça única, e a razão de a heurística
    /// antiga mandar o logotipo para o centro.
    /// </remarks>
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
    /// Rodapé de três parágrafos centralizados, como o do modelo de manual.
    /// </summary>
    /// <remarks>
    /// Endereço, autoria e data, um por parágrafo. Emendados numa linha só, os
    /// três viravam uma frase que atravessava a folha.
    /// </remarks>
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
    /// Rodapé com número de página: texto, tabulação e o campo `PAGE`.
    /// </summary>
    /// <remarks>
    /// As três coisas que a faixa mostra e que não são a mesma coisa por
    /// dentro. Só a primeira tem `w:t` onde escrever; a tabulação vira um
    /// espaço na tela mas continua sendo `w:tab` no arquivo, e o campo é
    /// calculado a cada abertura.
    /// </remarks>
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

    /// <summary>
    /// Quebra de página **sozinha** num parágrafo, sem mais nada.
    /// </summary>
    public static byte[] WithLonePageBreak() => Build((body, _) =>
    {
        body.AppendChild(Paragraph("Primeira página."));
        body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        body.AppendChild(Paragraph("Segunda página."));
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

    /// <summary>
    /// Cabeçalho em grade, como o cabeçalho corporativo do corpus.
    /// </summary>
    /// <remarks>
    /// Quatro colunas de grade e três linhas. O logotipo mora na primeira
    /// coluna, mesclada verticalmente pelas três; à direita, duas colunas da
    /// grade viram uma só por `w:gridSpan`. É a estrutura que, achatada em
    /// esquerda-centro-direita, virava uma fileira de palavras por cima da
    /// primeira linha do texto.
    /// </remarks>
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

            // A linha em que a borda de baixo é apagada na célula: é assim que
            // duas linhas do arquivo viram uma moldura só na tela.
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
    /// Documento que pede uma fonte que a máquina pode não ter.
    /// </summary>
    /// <remarks>
    /// `word/fontTable.xml` diz de que tipo cada fonte é. Sem consultá-lo, uma
    /// fonte ausente cai na próxima da pilha do editor, que termina em serifa —
    /// e a capa do modelo de manual, que pede Segoe UI, saía com o título em
    /// Times enquanto o LibreOffice o desenha sem serifa.
    /// </remarks>
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
    /// Cabeçalho que é um **grupo de formas**: logotipo e caixa de título.
    /// </summary>
    /// <remarks>
    /// O idioma do cabeçalho corporativo. A âncora diz onde o grupo está e que
    /// tamanho ele tem; `a:chOff`/`a:chExt` dizem em que régua as coordenadas de
    /// dentro foram escritas. Sem desembrulhar, cada peça recebia a caixa do
    /// grupo inteiro — o logotipo era esticado para a faixa toda — e as caixas
    /// de texto não saíam de lugar nenhum, porque só se procurava imagem.
    /// </remarks>
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
                              <w:p><w:r><w:rPr><w:b/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
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
            AnchoredDrawing(
                part.GetIdOfPart(image),
                cx: 10287325,
                cy: 2885145,
                rotation: 16200000,
                placed: true))));
    });

    /// <summary>
    /// A marca da capa e a quebra de página no mesmo parágrafo.
    /// </summary>
    /// <remarks>
    /// É como o modelo de manual encerra a capa: um parágrafo que não tem texto
    /// nenhum, só o desenho posicionado e o `w:br` que abre a folha seguinte.
    /// </remarks>
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
    /// Imagem ancorada que o LibreOffice grava no lugar do próprio parágrafo.
    /// </summary>
    /// <remarks>
    /// `wp:anchor` sem deslocamento, centralizada na coluna e com a largura
    /// dela. É como um documento de capturas de tela é escrito inteiro — e
    /// tratá-la como posição na folha fazia trinta imagens deixarem de ocupar
    /// altura, o texto se fechar por cima delas e um documento de doze folhas
    /// virar quatro.
    /// </remarks>
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

    /// <summary>
    /// A medida da linha: estilo padrão, marca de parágrafo e entrelinha travada.
    /// </summary>
    /// <remarks>
    /// Três coisas que o corpus real faz em todo parágrafo e o leitor ignorava:
    ///
    /// - o texto do corpo não declara `w:pStyle` e mora no estilo marcado
    ///   `w:default="1"`, que é onde estão a fonte e o corpo dele;
    /// - a fonte da linha vem da **marca de parágrafo** (`w:pPr/w:rPr`), que é
    ///   com o que o Word mede a linha e dá altura ao parágrafo vazio;
    /// - quem não diz nada sobre entrelinha está pedindo o espaçamento simples,
    ///   e não o padrão de quem abre o arquivo.
    /// </remarks>
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

        // Entrelinha travada em 9 pt — `exact` voltava nula e virava a do editor.
        var exact = Paragraph("Entrelinha travada.");
        exact.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "180", LineRule = LineSpacingRuleValues.Exact });
        body.AppendChild(exact);

        // Uma vez e meia, que é o outro valor que aparece na prática.
        var loose = Paragraph("Entrelinha de uma vez e meia.");
        loose.ParagraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto });
        body.AppendChild(loose);

        // 271/240 em Arial: o múltiplo mais comum do corpus, na fonte mais
        // comum dele.
        var multiple = Paragraph("Arial e um pouco mais de linha.");
        multiple.ParagraphProperties = new ParagraphProperties(
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Arial" }),
            new SpacingBetweenLines { Line = "271", LineRule = LineSpacingRuleValues.Auto });
        body.AppendChild(multiple);

        // Fonte que o instalador não leva: a substituta depende da máquina.
        var unknown = Paragraph("Numa fonte que ninguém tem.");
        unknown.ParagraphProperties = new ParagraphProperties(
            new ParagraphMarkRunProperties(new RunFonts { Ascii = "Fonte Fantasma" }),
            new SpacingBetweenLines { Line = "271", LineRule = LineSpacingRuleValues.Auto });
        body.AppendChild(unknown);
    });

    /// <summary>Um parágrafo que pede para ficar com o seguinte, outro que não.</summary>
    /// <summary>
    /// Estilo com entrelinha e recuo; parágrafo que redeclara **só o espaço**.
    /// </summary>
    /// <remarks>
    /// A forma exata do documento de evidências do corpus, e a que expôs o
    /// defeito: o estilo `BodyText` pede entrelinha 276 e recuo, e cada
    /// parágrafo redeclara apenas `w:before` e `w:after`. Substituindo o
    /// `w:spacing` inteiro, a entrelinha do estilo sumia — e cada linha saía
    /// 1,15 vez mais curta do que no LibreOffice.
    /// </remarks>
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
    /// Marca de seção no meio do texto, como o LibreOffice a grava.
    /// </summary>
    /// <remarks>
    /// A seção termina num `w:sectPr` guardado dentro do `w:pPr` de um
    /// parágrafo vazio: o parágrafo **é** a marca. O documento de evidências do
    /// corpus tem sete seções de mesma geometria e seis marcas espalhadas pelo
    /// meio do texto.
    /// </remarks>
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

        // A marca do Word vem da área de uso privado do Unicode — é assim que
        // ele grava os glifos de Symbol e Wingdings. Aqui, o quadrado.
        var level = new Level(
            new NumberingFormat { Val = NumberFormatValues.Bullet },
            new LevelText { Val = "\uF0A7" },
            new PreviousParagraphProperties(new Indentation { Left = "720", Hanging = "360" }))
        {
            LevelIndex = 0,
        };

        numbering.Numbering = new Numbering(
            new AbstractNum(level) { AbstractNumberId = 1 },
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

    /// <param name="placed">
    /// Fora do fluxo, com posição de verdade: a marca da capa, que sai para a
    /// margem e deixa o texto passar por baixo. Com `false`, a forma que o
    /// LibreOffice usa para "imagem no próprio parágrafo" — ancorada, mas no
    /// lugar em que o fluxo já a poria.
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

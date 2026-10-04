using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// O documento novo salvo como <c>.docx</c>: o pacote de <see cref="DocxTemplate"/>
/// faz o papel de original, e o defeito dele todo documento novo herdaria.
/// </summary>
public class DocxTemplateTests
{
    private static PageSetupDto A4() =>
        new("A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null);

    private static Node Text(string type, string text, params (string Name, JsonNode? Value)[] attrs)
    {
        var node = Node.Of(type);
        node.Content = [new Node { Type = "text", Text = text }];
        foreach (var (name, value) in attrs) node.With(name, value);
        return node;
    }

    [Fact]
    public void PacoteNovoPassaPeloValidadorSemErro()
    {
        Roundtrip.AssertSchema(DocxTemplate.Create(A4()));
    }

    [Fact]
    public void PacoteNovoTemAsPartesMinimasENaoTemNumeracao()
    {
        // A numeração nasce com a primeira lista (NumberingFactory).
        var parts = Roundtrip.PartsOf(DocxTemplate.Create(A4())).Keys.Order(StringComparer.Ordinal);

        Assert.Equal(
            [
                "[Content_Types].xml", "_rels/.rels", "docProps/app.xml", "docProps/core.xml",
                "word/_rels/document.xml.rels", "word/document.xml", "word/fontTable.xml",
                "word/settings.xml", "word/styles.xml",
            ],
            parts);
    }

    [Fact]
    public void PacoteNovoEDeterministico()
    {
        // O mesmo pedido dá os mesmos bytes.
        Assert.Equal(DocxTemplate.Create(A4()), DocxTemplate.Create(A4()));
    }

    [Fact]
    public void PaginaDoPacoteVemDaConfiguracao()
    {
        var page = new PageSetupDto("Letter", "landscape", new MarginsDto(10, 20, 30, 40), null, null);
        var xml = Roundtrip.XmlOf(DocxTemplate.Create(page));

        Assert.Contains("<w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"", xml, StringComparison.Ordinal);
        // 10 mm são 567 twips; 40 mm, 2268.
        Assert.Contains("w:top=\"567\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:left=\"2268\"", xml, StringComparison.Ordinal);

        var reopened = Roundtrip.Open(DocxTemplate.Create(page)).Page;
        Assert.Equal("Letter", reopened.Size);
        Assert.Equal("landscape", reopened.Orientation);
    }

    [Fact]
    public void PacoteNovoDeclaraEstilosConfiguracaoEPropriedades()
    {
        var bytes = DocxTemplate.Create(A4());
        var styles = Roundtrip.XmlOf(bytes, "word/styles.xml");

        // Sem tema no pacote, a fonte vai por nome.
        Assert.Contains("w:ascii=\"Times New Roman\"", styles, StringComparison.Ordinal);
        Assert.DoesNotContain("asciiTheme", styles, StringComparison.Ordinal);

        for (var level = 1; level <= 6; level++)
        {
            Assert.Contains($"w:styleId=\"Heading{level}\"", styles, StringComparison.Ordinal);
            Assert.Contains($"<w:name w:val=\"heading {level}\"", styles, StringComparison.Ordinal);
            Assert.Contains($"<w:outlineLvl w:val=\"{level - 1}\"", styles, StringComparison.Ordinal);
        }

        foreach (var id in (string[])
                 ["Normal", "DefaultParagraphFont", "TableNormal", "NoList", "TableGrid", "ListParagraph", "Hyperlink"])
        {
            Assert.Contains($"w:styleId=\"{id}\"", styles, StringComparison.Ordinal);
        }

        var settings = Roundtrip.XmlOf(bytes, "word/settings.xml");
        Assert.Contains("w:name=\"compatibilityMode\"", settings, StringComparison.Ordinal);
        Assert.Contains("w:val=\"15\"", settings, StringComparison.Ordinal);
        Assert.Contains("w:defaultTabStop", settings, StringComparison.Ordinal);

        Assert.Matches("<(\\w+:)?Application>Librevia</(\\w+:)?Application>", Roundtrip.XmlOf(bytes, "docProps/app.xml"));
        Assert.Contains("<dc:creator></dc:creator>", Roundtrip.XmlOf(bytes, "docProps/core.xml"),
            StringComparison.Ordinal);
    }

    [Fact]
    public void EstilosDoPacoteReproduzemATelaDoDocumentoNovo()
    {
        // Sem folha de estilos, o pacote leva a tabela de BuiltinStyles; reaberto,
        // tem de dizer o mesmo, senão a paginação muda.
        var model = new DocumentModelDto(A4(), Node.Of("doc",
            Text("heading", "Título", ("level", 1)),
            Text("paragraph", "Corpo do texto.")));

        var (saved, _) = Roundtrip.Save(DocxTemplate.Create(A4()), model);
        // Achatado: o que o arquivo vale, estilo incluído.
        var blocks = Roundtrip.OpenFlat(saved).Doc.Content!;

        var heading = blocks[0];
        Assert.Equal("heading", heading.Type);
        Assert.Equal("22pt", heading.Attrs!["fontSize"]!.GetValue<string>());
        Assert.Equal(22, heading.Attrs["spaceBefore"]!.GetValue<double>());

        var paragraph = blocks[1];
        Assert.Equal("12pt", paragraph.Attrs!["fontSize"]!.GetValue<string>());
        Assert.StartsWith("Times New Roman", paragraph.Attrs["fontFamily"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal(7.2, paragraph.Attrs["spaceBefore"]!.GetValue<double>());
        Assert.Equal(12, paragraph.Attrs["spaceAfter"]!.GetValue<double>());
        // O leitor arredonda 313/240 a 1,30 antes de multiplicar por 1,1499: 1,4949.
        Assert.InRange(double.Parse(paragraph.Attrs["lineHeight"]!.GetValue<string>(),
            System.Globalization.CultureInfo.InvariantCulture), 1.49, 1.51);
    }

    [Fact]
    public void DocumentoNovoVaiEVoltaComTituloListaTabelaEImagem()
    {
        var cell = (string text) => Node.Of("tableCell", Text("paragraph", text));
        var table = Node.Of("table", Node.Of("tableRow", cell("A1"), cell("B1")), Node.Of("tableRow", cell("A2"), cell("B2")));
        var image = Node.Of("image")
            .With("src", "data:image/png;base64," + Convert.ToBase64String(Fixtures.SquarePng()));

        var model = new DocumentModelDto(A4(), Node.Of("doc",
            Text("heading", "Relatório", ("level", 1)),
            Text("paragraph", "Primeiro parágrafo."),
            Node.Of("bulletList",
                Node.Of("listItem", Text("paragraph", "item um")),
                Node.Of("listItem", Text("paragraph", "item dois"))),
            table,
            image));

        var (saved, result) = Roundtrip.Save(DocxTemplate.Create(A4()), model);

        // Nenhum bloco tem `oid`: tudo é novo.
        Assert.Equal(0, result.PreservedBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("word/numbering.xml", Roundtrip.PartsOf(saved).Keys);
        Assert.Contains("<w:pStyle w:val=\"Heading1\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);

        var reopened = Roundtrip.Open(saved);
        var nodes = Roundtrip.Walk(reopened.Doc).ToList();

        var heading = Assert.Single(nodes, node => node.Type == "heading");
        Assert.Equal(1, heading.Attrs!["level"]!.GetValue<int>());
        Assert.Equal(2, nodes.Count(node => node.Type == "listItem"));
        Assert.Contains(nodes, node => node.Type == "bulletList");
        Assert.Equal(4, nodes.Count(node => node.Type == "tableCell"));
        Assert.Single(nodes, node => node.Type == "image");
        Assert.Equal("RelatórioPrimeiro parágrafo.item umitem doisA1B1A2B2", Roundtrip.TextOf(reopened));
    }

    [Fact]
    public void SegundaGravacaoSobreOPacoteGravadoFunciona()
    {
        // Depois da primeira gravação, os bytes gravados viram o original.
        var first = new DocumentModelDto(A4(), Node.Of("doc", Text("paragraph", "Versão um.")));
        var (saved, _) = Roundtrip.Save(DocxTemplate.Create(A4()), first);

        var second = new DocumentModelDto(A4(), Node.Of("doc",
            Text("heading", "Título novo", ("level", 2)),
            Text("paragraph", "Versão dois.")));
        var (again, _) = Roundtrip.Save(saved, second);

        Assert.Equal("Título novoVersão dois.", Roundtrip.TextOf(Roundtrip.Open(again)));
    }

    [Fact]
    public void CabecalhoDeTextoSimplesDoDocumentoNovoChegaAoArquivo()
    {
        // Cabeçalho e rodapé de texto simples, com `{n}` no lugar do número.
        var page = A4() with { HeaderText = "Relatório anual", FooterText = "Página {n} de {total}" };
        var model = new DocumentModelDto(page, Node.Of("doc", Text("paragraph", "Corpo.")));

        var (saved, result) = Roundtrip.Save(DocxTemplate.Create(page), model);
        Assert.Empty(result.Inventory.Lost);

        var parts = Roundtrip.PartsOf(saved).Keys.ToList();
        var header = Assert.Single(parts, name => name.StartsWith("word/header", StringComparison.Ordinal));
        var footer = Assert.Single(parts, name => name.StartsWith("word/footer", StringComparison.Ordinal));
        Assert.Contains("Relatório anual", Roundtrip.XmlOf(saved, header), StringComparison.Ordinal);
        Assert.Contains("NUMPAGES", Roundtrip.XmlOf(saved, footer), StringComparison.Ordinal);

        // Texto mudado depois da primeira gravação: a parte é reescrita.
        var edited = model with { Page = page with { HeaderText = "Relatório revisto" } };
        var (again, _) = Roundtrip.Save(saved, edited);
        Assert.Contains("Relatório revisto", Roundtrip.XmlOf(again, header), StringComparison.Ordinal);
    }

    [Fact]
    public void FaixaDeOutroPacoteNaoDerrubaAGravacaoEEntraNoInventario()
    {
        // O `.sdoc` que foi `.docx` guarda a faixa com ids de relação do pacote de
        // origem (`rId5:0:0`), que o pacote mínimo não tem: a faixa se perde, com
        // aviso, e o arquivo sai.
        var band = new BandDto(
            [new PieceDto("text", "Cabeçalho do arquivo de origem", Pid: "rId5:0:0")],
            [],
            [],
            true,
            // A caixa de texto do cabeçalho corporativo, pelo endereço `rId5#0`.
            [new FloatDto(
                "text", null, [], 0, 0, 0, "column", 0, null, "paragraph", 0, null, false, "none",
                BoxId: "rId5#0")],
            // As peças da grade do cabeçalho de evidências moram nas células.
            [new BandRowDto([new BandCellDto(
                [new PieceDto("text", "Chamado 10001", Pid: "rId5:1:0")], 1, 1, 1, null, "tlbr")])]);

        var page = A4() with { Header = band };
        var model = new DocumentModelDto(page, Node.Of("doc", Text("paragraph", "Primeira linha do corpo.")));

        var (saved, result) = Roundtrip.Save(DocxTemplate.Create(A4()), model);

        Assert.Contains("cabeçalho e rodapé do arquivo .docx de origem", result.Inventory.Lost);
        Assert.Contains(
            "Primeira linha do corpo.",
            Roundtrip.XmlOf(saved, "word/document.xml"),
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            Roundtrip.PartsOf(saved).Keys,
            name => name.StartsWith("word/header", StringComparison.Ordinal));
    }

    [Fact]
    public void TituloNovoNumDocxComTtulo1SaiComoTtulo1()
    {
        // No Word em português o id é `Ttulo1`, e o nome continua `heading 1`.
        var original = WithLocalizedHeadingStyle();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var paragraph = model.Doc.Content![0];
        model.Doc.Content[0] = Text("heading", "Virou título", ("level", 1));
        Assert.Equal("paragraph", paragraph.Type);

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.Contains("<w:pStyle w:val=\"Ttulo1\"", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.DoesNotContain("Heading1", Roundtrip.XmlOf(saved), StringComparison.Ordinal);

        // O estilo já existia: os estilos do documento não são tocados.
        Assert.Equal(Roundtrip.PartsOf(original)["word/styles.xml"], Roundtrip.PartsOf(saved)["word/styles.xml"]);

        var heading = Roundtrip.Open(saved).Doc.Content![0];
        Assert.Equal("heading", heading.Type);
        Assert.Equal(1, heading.Attrs!["level"]!.GetValue<int>());
    }

    [Fact]
    public void TituloNovoSemEstiloNoPacoteLevaADefinicaoDoModelo()
    {
        // O pacote não define título: a definição vem do modelo do documento novo, e
        // `word/styles.xml` vira parte tocada.
        var original = WithLocalizedHeadingStyle();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        model.Doc.Content![0] = Text("heading", "Subtítulo", ("level", 2));

        var (saved, _) = Roundtrip.Save(original, model);
        var styles = Roundtrip.XmlOf(saved, "word/styles.xml");

        Assert.Contains("w:styleId=\"Heading2\"", styles, StringComparison.Ordinal);
        Assert.Contains("<w:name w:val=\"heading 2\"", styles, StringComparison.Ordinal);
        Assert.Contains("w:styleId=\"Ttulo1\"", styles, StringComparison.Ordinal);

        var heading = Roundtrip.OpenFlat(saved).Doc.Content![0];
        Assert.Equal("heading", heading.Type);
        Assert.Equal(2, heading.Attrs!["level"]!.GetValue<int>());
        Assert.Equal("17pt", heading.Attrs["fontSize"]!.GetValue<string>());
    }

    [Fact]
    public void TituloDeOutroPacoteNaoApontaEstiloQueEstePacoteNaoDefine()
    {
        // O `.sdoc` de um `.docx` em português leva `Ttulo1`, que o pacote mínimo não define.
        var model = new DocumentModelDto(A4(), Node.Of("doc",
            Text("heading", "Título vindo de fora", ("level", 1), ("styleId", "Ttulo1"))));

        var (saved, result) = Roundtrip.Save(DocxTemplate.Create(A4()), model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.DoesNotContain("w:val=\"Ttulo1\"", xml, StringComparison.Ordinal);
        Assert.Contains("<w:pStyle w:val=\"Heading1\"", xml, StringComparison.Ordinal);
        Assert.Empty(result.Inventory.Lost);

        var heading = Roundtrip.Open(saved).Doc.Content![0];
        Assert.Equal("heading", heading.Type);
        Assert.Equal(1, heading.Attrs!["level"]!.GetValue<int>());
    }

    [Fact]
    public void TituloRebaixadoPerdeOEstiloReconhecidoSoPeloNome()
    {
        // `Überschrift1` se reconhece pelo nome: tela e arquivo dizem o mesmo ao rebaixá-lo.
        var original = WithHeadingNamedOnly();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var heading = model.Doc.Content![0];
        Assert.Equal("heading", heading.Type);

        var lowered = Text("paragraph", "Virou parágrafo", ("styleId", "Überschrift1"));
        lowered.Attrs!["oid"] = heading.Attrs!["oid"]!.DeepClone();
        model.Doc.Content[0] = lowered;

        var (saved, _) = Roundtrip.Save(original, model);

        Assert.DoesNotContain("Überschrift1", Roundtrip.XmlOf(saved), StringComparison.Ordinal);
        Assert.Equal("paragraph", Roundtrip.Open(saved).Doc.Content![0].Type);
    }

    [Fact]
    public void DocxComTituloLocalizadoPeloNomeAbreESalvaSemReescreverNada()
    {
        // O título reconhecido pelo nome não pode fazer o escritor discordar do leitor.
        var original = WithHeadingNamedOnly();
        var model = Roundtrip.Clone(Roundtrip.Open(original));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Equal(2, result.PreservedBlocks);
        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Equal(Roundtrip.PartsOf(original)["word/styles.xml"], Roundtrip.PartsOf(saved)["word/styles.xml"]);

        var reopened = Roundtrip.Open(saved);
        Assert.Equal("heading", reopened.Doc.Content![0].Type);
        Assert.Equal(1, reopened.Doc.Content[0].Attrs!["level"]!.GetValue<int>());
        Assert.Equal("Título alemãoUm parágrafo.", Roundtrip.TextOf(reopened));
    }

    /// <summary>Um DOCX como o Word em português grava: `Ttulo1`, de nome `heading 1`.</summary>
    internal static byte[] WithLocalizedHeadingStyle()
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var part = document.AddMainDocumentPart();
            var styles = part.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(
                    new StyleName { Val = "heading 1" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new OutlineLevel { Val = 0 }),
                    new StyleRunProperties(new Bold(), new FontSize { Val = "32" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Ttulo1",
                });

            part.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Um parágrafo."))),
                new SectionProperties(
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U, Header = 708U, Footer = 708U, Gutter = 0U })));
        }

        return buffer.ToArray();
    }

    /// <summary>
    /// Título que só se reconhece pelo <c>w:name</c>: <c>Überschrift1</c> é o id do
    /// Word em alemão, e o nome é <c>heading 1</c> em qualquer idioma.
    /// </summary>
    private static byte[] WithHeadingNamedOnly()
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var part = document.AddMainDocumentPart();
            var styles = part.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Style(
                    new StyleName { Val = "heading 1" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new OutlineLevel { Val = 0 }),
                    new StyleRunProperties(new Bold(), new FontSize { Val = "32" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Überschrift1",
                });

            part.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "Überschrift1" }),
                    new Run(new Text("Título alemão"))),
                new Paragraph(new Run(new Text("Um parágrafo."))),
                new SectionProperties(
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U, Header = 708U, Footer = 708U, Gutter = 0U })));
        }

        return buffer.ToArray();
    }
}

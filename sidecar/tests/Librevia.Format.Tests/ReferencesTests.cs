using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Referências (M8): marcadores, links internos, campos e sumário.
/// </summary>
/// <remarks>
/// O documento é o de <see cref="Fixtures.WithReferences"/>, que tem a forma exata
/// do Word. A primeira promessa é a de sempre — abrir e gravar sem editar não
/// reescreve nada, e cada elemento do corpo volta como estava —; as outras são o
/// que o editor passa a saber fazer com o que antes só atravessava intacto.
/// </remarks>
public class ReferencesTests
{
    private static string? AttrOf(Node node, string name) =>
        node.Attrs?.GetValueOrDefault(name) is { } value &&
        value.GetValueKind() == System.Text.Json.JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static List<Node> NodesOf(DocumentModelDto model, string type) =>
        [.. Roundtrip.Walk(model.Doc).Where(node => node.Type == type)];

    private static List<string> BodyXml(byte[] docx)
    {
        using var stream = new MemoryStream(docx.ToArray());
        using var document = WordprocessingDocument.Open(stream, false);
        return [.. document.MainDocumentPart!.Document!.Body!.ChildElements.Select(element => element.OuterXml)];
    }

    [Fact]
    public void AbrirEGravarSemEditarDevolveCadaElementoDoCorpoComoEstava()
    {
        var original = Fixtures.WithReferences();
        var (saved, result) = Roundtrip.Save(original, Roundtrip.Clone(Roundtrip.Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        // O corpo, elemento por elemento — o `w:bookmarkEnd` solto entre dois
        // parágrafos incluído, que antes caía fora em silêncio.
        Assert.Equal(BodyXml(original), BodyXml(saved));

        var before = Roundtrip.PartsOf(original);
        var after = Roundtrip.PartsOf(saved);
        Assert.DoesNotContain(
            before.Keys.Where(name => name != "word/document.xml"),
            name => !before[name].SequenceEqual(after[name]));
    }

    [Fact]
    public void OsMarcadoresViramNosComNomeEId()
    {
        var model = Roundtrip.Open(Fixtures.WithReferences());

        var starts = NodesOf(model, "bookmarkStart")
            .Select(node => (AttrOf(node, "name"), AttrOf(node, "bid")))
            .ToList();
        Assert.Contains(("_Toc100", "0"), starts);
        Assert.Contains(("Resumo", "1"), starts);
        Assert.Contains(("_Ref200", "3"), starts);

        // A ponta que mora no corpo, entre dois parágrafos, não é nó: fica com o
        // bloco seguinte e volta antes dele.
        Assert.DoesNotContain(NodesOf(model, "bookmarkEnd"), node => AttrOf(node, "bid") == "1");
    }

    [Fact]
    public void OLinkInternoViraEnderecoComCerquilha()
    {
        var model = Roundtrip.Open(Fixtures.WithReferences());

        var hrefs = Roundtrip.Walk(model.Doc)
            .SelectMany(node => node.Marks ?? [])
            .Where(mark => mark.Type == "link")
            .Select(mark => (mark.Attrs?.GetValueOrDefault("href")?.GetValue<string>()))
            .ToList();
        Assert.Contains("#Resumo", hrefs);
    }

    [Fact]
    public void ParagrafoEditadoGravaOMarcadorEOLinkInternoDoModelo()
    {
        var original = Fixtures.WithReferences();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Assert.True(Roundtrip.EditFirstTextContaining(model, "Veja o ", ". Confira o "));

        // Um marcador novo no parágrafo editado, como o diálogo o insere.
        var edited = Roundtrip.Walk(model.Doc).First(node =>
            node.Content?.Any(child => child.Text == ". Confira o ") == true);
        edited.Content!.Insert(0, Node.Of("bookmarkStart").With("name", "Novo").With("bid", "9"));
        edited.Content!.Add(Node.Of("bookmarkEnd").With("bid", "9"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Matches("<w:bookmarkStart (w:id=\"9\" w:name=\"Novo\"|w:name=\"Novo\" w:id=\"9\") ?/>", xml);
        Assert.Contains("w:anchor=\"Resumo\"", xml, StringComparison.Ordinal);
        // O link interno não ganha relacionamento.
        Assert.DoesNotContain("relationships/hyperlink", Roundtrip.XmlOf(saved, "word/_rels/document.xml.rels"),
            StringComparison.Ordinal);

        var reopened = Roundtrip.Open(saved);
        Assert.Contains(NodesOf(reopened, "bookmarkStart"), node => AttrOf(node, "name") == "Novo");
    }

    [Fact]
    public void MarcadorApagadoDoParagrafoNaoVolta()
    {
        // A pessoa excluiu o marcador pelo diálogo: o parágrafo é reescrito sem
        // ele, e o escritor não o copia de volta do original como fazia antes.
        var original = Fixtures.WithReferences();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var heading = Roundtrip.Walk(model.Doc).First(node =>
            node.Content?.Any(child => AttrOf(child, "name") == "_Toc100") == true);
        heading.Content!.RemoveAll(child => child.Type is "bookmarkStart" or "bookmarkEnd");

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.DoesNotContain("w:name=\"_Toc100\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagrafoColadoNaoDuplicaMarcadorDoModelo()
    {
        var original = Fixtures.WithReferences();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var heading = model.Doc.Content!.First(node => node.Type == "heading");
        var copy = Roundtrip.Clone(model).Doc.Content!.First(node => node.Type == "heading");
        copy.Content!.Add(new Node { Type = "text", Text = " (cópia)" });
        model.Doc.Content!.Insert(model.Doc.Content.IndexOf(heading) + 1, copy);

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Single(Regex.Matches(xml, "w:name=\"_Toc100\""));
    }

    [Fact]
    public void RascunhoDeAntesDasReferenciasContinuaReconhecido()
    {
        // O rascunho gravado antes do M8 não tem marcador nem link interno nos
        // nós. Comparado com a leitura nova, todo parágrafo marcado seria
        // reescrito; com a leitura de então, nada muda.
        var original = Fixtures.WithReferences();
        var legacy = Roundtrip.Clone(new DocumentModelDto(
            Roundtrip.Open(original).Page,
            DocxReaderOf(original, references: false)));

        var (saved, result) = Roundtrip.Save(original, legacy with { BeforeReferences = true });

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(BodyXml(original), BodyXml(saved));
    }

    [Fact]
    public void MarcadorEntreLinhasDeTabelaEditadaContinuaNoLugar()
    {
        // O Word grava assim o marcador que abraça linhas inteiras: as pontas são
        // filhas de `w:tbl` e de `w:tr`, fora de qualquer parágrafo. Editar uma
        // célula reescreve a tabela, e as pontas têm de voltar onde estavam.
        const string cell = "<w:tc><w:tcPr><w:tcW w:w=\"4500\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{0}</w:t></w:r></w:p></w:tc>";
        var body =
            "<w:tbl><w:tblPr><w:tblW w:w=\"9000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4500\"/><w:gridCol w:w=\"4500\"/></w:tblGrid>" +
            "<w:tr>" + string.Format(cell, "A") + string.Format(cell, "B") + "</w:tr>" +
            "<w:bookmarkStart w:id=\"5\" w:name=\"Linhas\"/>" +
            "<w:tr><w:bookmarkStart w:id=\"6\" w:name=\"Celula\"/>" + string.Format(cell, "C") + "<w:bookmarkEnd w:id=\"6\"/>" + string.Format(cell, "D") + "</w:tr>" +
            "<w:bookmarkEnd w:id=\"5\"/></w:tbl><w:p/>";
        var original = Fixtures.BuildFromXml(body, "");
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Assert.True(Roundtrip.EditFirstTextContaining(model, "A", "A editado"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Matches("</w:tr><w:bookmarkStart w:name=\"Linhas\" w:id=\"5\" ?/><w:tr>", xml);
        Assert.Matches("<w:tr><w:bookmarkStart w:name=\"Celula\" w:id=\"6\" ?/><w:tc>", xml);
        Assert.Matches("</w:tr><w:bookmarkEnd w:id=\"5\" ?/></w:tbl>", xml);
    }

    // --- campos e sumário ----------------------------------------------------

    private static string Instruction(Node field) => AttrOf(field, "instr")!.Trim();

    [Fact]
    public void CamposViramNosComInstrucaoEResultado()
    {
        var opened = DocxReader.Read(Fixtures.WithReferences());
        var fields = NodesOf(opened.Model, "field");

        Assert.Contains(fields, field => Instruction(field) == "SEQ Figura \\* ARABIC" && AttrOf(field, "result") == "1");
        Assert.Contains(fields, field => Instruction(field) == "REF _Ref200 \\h" && AttrOf(field, "result") == "Figura 1");
        Assert.Contains(fields, field => Instruction(field) == "PAGEREF _Ref200 \\h");

        // Representados, os campos e o sumário deixam de travar o documento.
        Assert.Empty(opened.Inventory.Structural);
    }

    [Fact]
    public void OSumarioViraUmBlocoComAsEntradasDentro()
    {
        var model = Roundtrip.Open(Fixtures.WithReferences());
        var toc = Assert.Single(NodesOf(model, "tableOfContents"));

        Assert.StartsWith("TOC \\o \"1-3\"", Instruction(toc));
        Assert.Equal(1, toc.Attrs!["head"]!.GetValue<int>());
        Assert.True(toc.Attrs!["sdt"]!.GetValue<bool>());
        Assert.Equal(4, toc.Content!.Count);

        // A entrada: o link para o `_Toc` do título, com o PAGEREF dentro.
        var entry = toc.Content[1];
        var pageRef = Assert.Single(entry.Content!, node => node.Type == "field");
        Assert.Equal("PAGEREF _Toc100 \\h", Instruction(pageRef));
        Assert.Contains(pageRef.Marks!, mark => mark.Type == "link");
        // As peças do campo TOC saíram do texto: nenhum `field` com TOC dentro.
        Assert.DoesNotContain(NodesOf(model, "field"), field => Instruction(field).StartsWith("TOC", StringComparison.Ordinal));
    }

    [Fact]
    public void SumarioEditadoVoltaComoCampoDentroDoControleDeConteudo()
    {
        var original = Fixtures.WithReferences();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Assert.True(Roundtrip.EditFirstTextContaining(model, "Escopo", "Escopo revisto"));

        var (saved, result) = Roundtrip.Save(original, model);
        var xml = Roundtrip.XmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("w:docPartGallery w:val=\"Table of Contents\"", xml, StringComparison.Ordinal);
        Assert.Single(Regex.Matches(xml, "> TOC \\\\o"));
        Assert.Equal(2, Regex.Matches(xml, "PAGEREF _Toc10").Count);

        var reopened = Roundtrip.Open(saved);
        var toc = Assert.Single(NodesOf(reopened, "tableOfContents"));
        Assert.Equal(4, toc.Content!.Count);
        Assert.Contains("Escopo revisto", Roundtrip.TextOf(reopened), StringComparison.Ordinal);
    }

    [Fact]
    public void ParagrafoComCampoEditadoMantemOCampo()
    {
        var original = Fixtures.WithReferences();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        Assert.True(Roundtrip.EditFirstTextContaining(model, " — Arquitetura", " — Arquitetura geral"));

        var (saved, result) = Roundtrip.Save(original, model);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        var reopened = Roundtrip.Open(saved);
        Assert.Contains(NodesOf(reopened, "field"), field => Instruction(field) == "SEQ Figura \\* ARABIC");
        // O marcador `_Ref` da legenda continua em volta do número.
        Assert.Contains(NodesOf(reopened, "bookmarkStart"), node => AttrOf(node, "name") == "_Ref200");
    }

    [Fact]
    public void SumarioSemControleDeConteudoEhUmBlocoDeVariosParagrafos()
    {
        var body = Fixtures.ReferencesBody;
        var inner = body[(body.IndexOf("<w:sdtContent>", StringComparison.Ordinal) + "<w:sdtContent>".Length)..body.IndexOf("</w:sdtContent>", StringComparison.Ordinal)];
        // Sem o título, que fora do controle de conteúdo seria um parágrafo comum.
        inner = inner[(inner.IndexOf("</w:p>", StringComparison.Ordinal) + "</w:p>".Length)..];
        var bare = inner + body[(body.IndexOf("</w:sdt>", StringComparison.Ordinal) + "</w:sdt>".Length)..];
        var original = Fixtures.BuildFromXml(bare, Fixtures.ReferencesStyles);

        var model = Roundtrip.Open(original);
        var toc = Assert.Single(NodesOf(model, "tableOfContents"));
        Assert.False(toc.Attrs!["sdt"]!.GetValue<bool>());
        Assert.Equal(3, toc.Content!.Count);

        var (unchanged, untouched) = Roundtrip.Save(original, Roundtrip.Clone(model));
        Assert.Equal(0, untouched.RewrittenBlocks);
        Assert.Equal(BodyXml(original), BodyXml(unchanged));

        var edited = Roundtrip.Clone(model);
        Assert.True(Roundtrip.EditFirstTextContaining(edited, "Introdução", "Introdução geral"));
        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, edited).Bytes);
        Assert.DoesNotContain("<w:sdt>", xml, StringComparison.Ordinal);
        Assert.Single(Regex.Matches(xml, "> TOC \\\\o"));
    }

    [Fact]
    public void SumarioNovoGanhaControleDeConteudoEParadaDeTabulacao()
    {
        // O que o editor manda ao inserir um sumário: o nó sem `oid`, com as
        // entradas já montadas.
        var original = Fixtures.WithReferences();
        var model = Roundtrip.Clone(Roundtrip.Open(original));
        var entry = Node.Of("paragraph").With("styleId", "Sumrio1");
        entry.Content =
        [
            new Node { Type = "text", Text = "Introdução\t", Marks = [Mark.Of("link", "href", "#_Toc100")] },
            new Node
            {
                Type = "field",
                Marks = [Mark.Of("link", "href", "#_Toc100")],
                Attrs = new() { ["instr"] = " PAGEREF _Toc100 \\h ", ["result"] = "2" },
            },
        ];
        var toc = Node.Of("tableOfContents", entry).With("instr", " TOC \\o \"1-2\" \\h ").With("head", 0).With("sdt", true);
        model.Doc.Content!.Add(toc);

        var xml = Roundtrip.XmlOf(Roundtrip.Save(original, model).Bytes);

        Assert.Equal(2, Regex.Matches(xml, "Table of Contents").Count);
        Assert.Contains("> TOC \\o \"1-2\" \\h <", xml, StringComparison.Ordinal);
        Assert.Matches("w:leader=\"dot\"[^>]*w:pos=\"\\d+\"|w:pos=\"\\d+\"[^>]*w:leader=\"dot\"", xml);
    }

    [Fact]
    public void CampoAninhadoContinuaTravandoEPerdeAoEditar()
    {
        // O campo dentro do campo não cabe no nó: fica como era — o resultado na
        // tela, o documento travado e o aviso se o parágrafo for reescrito.
        const string nested =
            "<w:p><w:r><w:t xml:space=\"preserve\">Antes </w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText xml:space=\"preserve\"> IF </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>1</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> = 1 \"um\" \"outro\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>um</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        var original = Fixtures.BuildFromXml(nested, "");
        var opened = DocxReader.Read(original);
        Assert.Contains(Inventory.Fields, opened.Inventory.Structural);

        var model = Roundtrip.Clone(opened.Model);
        Assert.True(Roundtrip.EditFirstTextContaining(model, "Antes ", "Depois "));
        var (_, result) = Roundtrip.Save(original, model);
        Assert.Contains("campo calculado num parágrafo que você editou", result.Inventory.Lost);
    }

    private static Node DocxReaderOf(byte[] bytes, bool references)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var part = document.MainDocumentPart!;
        var (content, _) = new BodyReader(part, new Inventory(), references: references)
            .Read(part.Document!.Body!);
        var doc = Node.Of("doc");
        doc.Content = content;
        return doc;
    }
}

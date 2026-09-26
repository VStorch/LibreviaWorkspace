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

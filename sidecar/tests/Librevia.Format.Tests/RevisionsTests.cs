using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// <summary>
/// Controle de alterações: lido, mostrado, devolvido ao arquivo.
/// </summary>
/// <remarks>
/// A revisão de texto é marca do editor (`insertion`/`deletion`), a da marca de
/// parágrafo e a da linha são atributos do bloco. Editar o parágrafo revisado
/// devolve o `w:ins` e o `w:del` com os mesmos ids — e por isso a revisão não
/// trava o documento.
/// </remarks>
public class RevisionsTests
{
    private const string Date = "2026-03-01T10:00:00Z";

    /// <summary>Um pouco de cada revisão que o editor lê.</summary>
    private static byte[] WithRevisions() => Fixtures.BuildFromXml(
        $"""
        <w:p><w:r><w:t>Intocado.</w:t></w:r></w:p>
        <w:p><w:r><w:t xml:space="preserve">Texto </w:t></w:r><w:ins w:id="10" w:author="Ana" w:date="{Date}"><w:r><w:t>inserido</w:t></w:r></w:ins><w:del w:id="11" w:author="Bruno" w:date="{Date}"><w:r><w:delText xml:space="preserve"> excluído</w:delText></w:r></w:del><w:ins w:id="12" w:author="Ana"><w:del w:id="13" w:author="Bruno"><w:r><w:delText xml:space="preserve"> ida e volta</w:delText></w:r></w:del></w:ins></w:p>
        <w:p><w:pPr><w:rPr><w:ins w:id="14" w:author="Ana" w:date="{Date}"/></w:rPr></w:pPr><w:r><w:t>Marca inserida</w:t></w:r></w:p>
        <w:p><w:moveFromRangeStart w:id="20" w:author="Ana" w:date="{Date}" w:name="move1"/><w:moveFrom w:id="21" w:author="Ana"><w:r><w:t>movido</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="20"/><w:r><w:t xml:space="preserve"> fica</w:t></w:r></w:p>
        <w:p><w:moveToRangeStart w:id="22" w:author="Ana" w:date="{Date}" w:name="move1"/><w:moveTo w:id="23" w:author="Ana"><w:r><w:t>movido</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="22"/></w:p>
        <w:p><w:hyperlink w:anchor="alvo"><w:ins w:id="15" w:author="Ana"><w:r><w:t>link inserido</w:t></w:r></w:ins></w:hyperlink></w:p>
        <w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="40" w:author="Ana"><w:rPr/></w:rPrChange></w:rPr><w:t>negrito revisado</w:t></w:r></w:p>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
        <w:tr><w:trPr><w:del w:id="30" w:author="Bruno"/></w:trPr><w:tc><w:p><w:r><w:t>Linha excluída</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:r><w:t>Linha que fica</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl><w:p/>
        """.ReplaceLineEndings(string.Empty),
        string.Empty);

    /// <summary>
    /// O XML com os atributos em ordem: o SDK grava os dele na ordem do esquema,
    /// e o fixture foi escrito à mão.
    /// </summary>
    private static string Canonical(string xml)
    {
        var document = System.Xml.Linq.XDocument.Parse(xml);
        foreach (var element in document.Descendants())
        {
            var attributes = element.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToList();
            element.RemoveAttributes();
            element.Add(attributes);
        }

        return document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static Node TextNode(DocumentModelDto model, string text) =>
        Walk(model.Doc).First(node => node.Type == "text" && node.Text == text);

    private static Mark? MarkOf(Node node, string type) => node.Marks?.FirstOrDefault(mark => mark.Type == type);

    private static string? AttrOf(Mark? mark, string name) => mark?.Attrs?.GetValueOrDefault(name)?.GetValue<string>();

    [Fact]
    public void LeInsercaoExclusaoEAsDuasJuntas()
    {
        var result = DocxReader.Read(WithRevisions());
        var model = result.Model;

        var inserted = MarkOf(TextNode(model, "inserido"), Revisions.Insertion);
        Assert.Equal("Ana", AttrOf(inserted, "author"));
        Assert.Equal(Date, AttrOf(inserted, "date"));
        Assert.Equal("10", AttrOf(inserted, "rid"));

        var deleted = MarkOf(TextNode(model, " excluído"), Revisions.Deletion);
        Assert.Equal("Bruno", AttrOf(deleted, "author"));
        Assert.Equal("11", AttrOf(deleted, "rid"));

        var both = TextNode(model, " ida e volta");
        Assert.Equal("12", AttrOf(MarkOf(both, Revisions.Insertion), "rid"));
        Assert.Equal("13", AttrOf(MarkOf(both, Revisions.Deletion), "rid"));

        // O link por fora da revisão.
        var link = TextNode(model, "link inserido");
        Assert.NotNull(MarkOf(link, "link"));
        Assert.Equal("15", AttrOf(MarkOf(link, Revisions.Insertion), "rid"));

        // Revisão de texto não trava nem avisa; a de formatação só avisa.
        Assert.Empty(result.Inventory.Structural);
        Assert.Equal([Inventory.FormatRevisions], result.Inventory.Invisible);
    }

    [Fact]
    public void LeMovimentacaoComoExclusaoEInsercao()
    {
        var model = Open(WithRevisions());
        var texts = Walk(model.Doc).Where(node => node.Type == "text" && node.Text == "movido").ToList();

        Assert.Equal(2, texts.Count);
        var from = MarkOf(texts[0], Revisions.Deletion);
        var to = MarkOf(texts[1], Revisions.Insertion);
        Assert.Equal("from", AttrOf(from, "move"));
        Assert.Equal("move1", AttrOf(from, "moveName"));
        Assert.Equal("to", AttrOf(to, "move"));
        Assert.Equal("move1", AttrOf(to, "moveName"));
    }

    [Fact]
    public void LeARevisaoDaMarcaDeParagrafoEDaLinha()
    {
        var model = Open(WithRevisions());

        var paragraph = model.Doc.Content!.First(block =>
            Walk(block).Any(node => node.Text == "Marca inserida"));
        var mark = (JsonObject)paragraph.Attrs!["markRevision"]!;
        Assert.Equal("ins", mark["kind"]!.GetValue<string>());
        Assert.Equal("14", mark["rid"]!.GetValue<string>());
        Assert.Equal(Date, mark["date"]!.GetValue<string>());

        var rows = model.Doc.Content!.First(block => block.Type == "table").Content!;
        var row = (JsonObject)rows[0].Attrs!["rowRevision"]!;
        Assert.Equal("del", row["kind"]!.GetValue<string>());
        Assert.Equal("30", row["rid"]!.GetValue<string>());
        Assert.Null(rows[1].Attrs?.GetValueOrDefault("rowRevision"));
    }

    [Fact]
    public void SalvarSemMexerNaoReescreveNadaEDevolveOXmlIgual()
    {
        var original = WithRevisions();
        var (saved, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Equal(Canonical(XmlOf(original)), Canonical(XmlOf(saved)));
    }

    [Fact]
    public void EditarOParagrafoRevisadoDevolveInsEDelComOsMesmosIds()
    {
        var original = WithRevisions();
        var model = Clone(Open(original));
        Assert.True(EditFirstTextContaining(model, "Texto ", "Texto mexido "));

        var (saved, result) = Save(original, model);
        var xml = XmlOf(saved);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Matches($"<w:ins w:author=\"Ana\" w:date=\"{Regex.Escape(Date)}\" w:id=\"10\"><w:r><w:t[^>]*>inserido</w:t>", xml);
        Assert.Matches("<w:del w:author=\"Bruno\"[^>]* w:id=\"11\"><w:r><w:delText xml:space=\"preserve\"> excluído</w:delText>", xml);
        Assert.Matches("<w:ins w:author=\"Ana\" w:id=\"12\"><w:del w:author=\"Bruno\" w:id=\"13\"><w:r><w:delText", xml);

        // E volta a abrir com as mesmas marcas.
        var reopened = Open(saved);
        Assert.Equal("10", AttrOf(MarkOf(TextNode(reopened, "inserido"), Revisions.Insertion), "rid"));
        Assert.Equal("11", AttrOf(MarkOf(TextNode(reopened, " excluído"), Revisions.Deletion), "rid"));
    }

    [Fact]
    public void ParagrafoPartidoGanhaIdNovoNaSegundaMetade()
    {
        var original = WithRevisions();
        var model = Clone(Open(original));
        var index = model.Doc.Content!.FindIndex(block => Walk(block).Any(node => node.Text == "inserido"));
        var copy = Clone(model with { Doc = new Node { Type = "doc", Content = [model.Doc.Content![index]] } });
        var duplicate = copy.Doc.Content![0];
        duplicate.Attrs?.Remove("oid");
        model.Doc.Content!.Insert(index + 1, duplicate);

        var xml = XmlOf(Save(original, model).Bytes);
        var ids = Regex.Matches(xml, "<w:(?:ins|del) [^>]*w:id=\"(\\d+)\"").Select(match => match.Groups[1].Value).ToList();

        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Contains("10", ids);
    }

    [Fact]
    public void RevisaoDeFormatacaoNoParagrafoEditadoEPerdaDeclarada()
    {
        var original = WithRevisions();
        var model = Clone(Open(original));
        Assert.True(EditFirstTextContaining(model, "negrito revisado", "negrito mexido"));

        var (_, result) = Save(original, model);

        Assert.Contains("revisão de formatação num parágrafo que você editou", result.Inventory.Lost);
    }

    [Fact]
    public void MovimentacaoEditadaViraExclusaoEInsercao()
    {
        var original = WithRevisions();
        var model = Clone(Open(original));
        Assert.True(EditFirstTextContaining(model, " fica", " fica mexido"));

        var (saved, result) = Save(original, model);
        var xml = XmlOf(saved);

        Assert.Contains(result.Inventory.Lost, loss => loss.StartsWith("movimentação de texto", StringComparison.Ordinal));
        Assert.DoesNotContain("moveFrom", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("moveTo", xml, StringComparison.Ordinal);
        Assert.Matches("<w:ins [^>]*w:id=\"23\"", xml);
    }

    [Fact]
    public void AceitarTudoSaiDentroDoEsquemaESemRevisao()
    {
        var original = WithRevisions();
        var model = Clone(Open(original));

        // O que "Aceitar tudo" faz no editor: a inserção fica, a exclusão sai, a
        // linha excluída sai, e a marca de parágrafo deixa de ser revisão.
        static void Accept(Node node)
        {
            node.Attrs?.Remove("markRevision");
            if (node.Content is not { } content) return;
            content.RemoveAll(child =>
                child.Marks?.Any(mark => mark.Type == Revisions.Deletion) == true ||
                (child.Attrs?.GetValueOrDefault("rowRevision") as JsonObject)?["kind"]?.GetValue<string>() == "del");
            foreach (var child in content)
            {
                child.Marks?.RemoveAll(mark => mark.Type == Revisions.Insertion);
                Accept(child);
            }
        }

        Accept(model.Doc);
        var (saved, _) = Save(original, model);
        var xml = XmlOf(saved);

        Assert.DoesNotMatch("<w:(ins|del|moveFrom|moveTo)[ >]", xml);
        Assert.DoesNotContain("excluído", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("Linha excluída", xml, StringComparison.Ordinal);
        Assert.Contains("inserido", xml, StringComparison.Ordinal);
    }

    // O que a leitura do rascunho anterior às revisões (`BeforeRevisions`) dá:
    // o inserido como texto comum, e nem o excluído nem a movimentação.
    private static void Legacy(Node node)
    {
        node.Attrs?.Remove("markRevision");
        node.Attrs?.Remove("rowRevision");
        if (node.Content is not { } content) return;
        content.RemoveAll(child =>
            child.Marks?.Any(mark => mark.Type == Revisions.Deletion || mark.Attrs?.ContainsKey("move") == true) == true);
        if (content.Count == 0) node.Content = null;
        foreach (var child in content)
        {
            child.Marks?.RemoveAll(mark => mark.Type == Revisions.Insertion);
            if (child.Marks?.Count == 0) child.Marks = null;
            Legacy(child);
        }
    }

    [Fact]
    public void RascunhoDeAntesDasRevisoesUsaALeituraDeEntao()
    {
        var original = WithRevisions();

        var model = Clone(Open(original));
        Legacy(model.Doc);
        var (_, result) = Save(original, model with { BeforeRevisions = true });

        Assert.Equal(0, result.RewrittenBlocks);
    }

    [Fact]
    public void RascunhoAntigoEditadoDeclaraAMovimentacaoEAFormatacao()
    {
        // O rascunho anterior às revisões (`BeforeRevisions`) não traz revisão
        // nenhuma: reescrever o parágrafo da movimentação ou o da formatação
        // revisada as perde.
        var original = WithRevisions();
        var model = Clone(Open(original));
        Legacy(model.Doc);
        Assert.True(EditFirstTextContaining(model, " fica", " fica mexido"));
        Assert.True(EditFirstTextContaining(model, "negrito revisado", "negrito mexido"));

        var (_, result) = Save(original, model with { BeforeRevisions = true });

        Assert.Contains(result.Inventory.Lost, loss => loss.StartsWith("marcas de revisão", StringComparison.Ordinal));
    }

    [Fact]
    public void CelulaInseridaContinuaTravando()
    {
        var inventory = DocxReader.Read(Fixtures.WithInsertedCell()).Inventory;

        Assert.Contains(Inventory.StructureRevisions, inventory.Structural);
    }

    [Fact]
    public void ControlarAlteracoesIdaEVolta()
    {
        var original = WithRevisions();
        Assert.Null(Open(original).TrackChanges);

        var (saved, _) = Save(original, Clone(Open(original)) with { TrackChanges = true });
        Assert.True(Open(saved).TrackChanges);
        Assert.Contains("trackRevisions", XmlOf(saved, "word/settings.xml"), StringComparison.Ordinal);

        var (off, _) = Save(saved, Clone(Open(saved)) with { TrackChanges = false });
        Assert.Null(Open(off).TrackChanges);
    }
}

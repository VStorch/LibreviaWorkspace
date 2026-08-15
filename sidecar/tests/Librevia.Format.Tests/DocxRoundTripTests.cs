using System.IO.Compression;
using System.Text.Json;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// A prova da Fase 4: editar um documento não pode custar o que não foi editado.
/// </summary>
public class DocxRoundTripTests
{
    private static DocumentModelDto Open(byte[] bytes) => DocxReader.Read(bytes).Model;

    private static (byte[] Bytes, SaveResult Result) Save(byte[] original, DocumentModelDto model) =>
        DocxWriter.Write(original, model);

    /// <summary>Clona pelo JSON — é como o modelo viaja de verdade.</summary>
    private static DocumentModelDto Clone(DocumentModelDto model) =>
        JsonSerializer.Deserialize<DocumentModelDto>(
            JsonSerializer.Serialize(model, DocxJson.Options), DocxJson.Options)!;

    private static IEnumerable<Node> Walk(Node node)
    {
        yield return node;
        foreach (var child in node.Content ?? [])
        {
            foreach (var deeper in Walk(child)) yield return deeper;
        }
    }

    private static string TextOf(DocumentModelDto model) =>
        string.Concat(Walk(model.Doc).Where(n => n.Type == "text").Select(n => n.Text));

    private static bool EditFirstTextContaining(DocumentModelDto model, string needle, string replacement)
    {
        var target = Walk(model.Doc)
            .FirstOrDefault(n => n.Type == "text" && n.Text?.Contains(needle, StringComparison.Ordinal) == true);
        if (target is null) return false;
        target.Text = replacement;
        return true;
    }

    private static Dictionary<string, byte[]> PartsOf(byte[] docx)
    {
        using var archive = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var stream = entry.Open();
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                return buffer.ToArray();
            },
            StringComparer.Ordinal);
    }

    // --- o critério de aceite ----------------------------------------------

    [Fact]
    public void EditingOneParagraphKeepsCommentsOnTheOthers()
    {
        // Este é *o* teste da Fase 4. Se ele falhar, a estratégia de edição
        // cirúrgica não está funcionando e voltamos ao problema do §6.1:
        // o usuário corrige uma vírgula e perde os comentários dos revisores.
        var original = Fixtures.WithComment();
        var model = Clone(Open(original));

        Assert.True(EditFirstTextContaining(model, "sem comentário", "Texto trocado."));

        var (saved, result) = Save(original, model);

        Assert.Contains("word/comments.xml", PartsOf(saved).Keys);
        Assert.Equal(PartsOf(original)["word/comments.xml"], PartsOf(saved)["word/comments.xml"]);
        Assert.Equal(1, result.RewrittenBlocks);
        Assert.True(result.PreservedBlocks >= 2);

        // A âncora precisa continuar no corpo, senão o comentário vira órfão.
        var body = System.Text.Encoding.UTF8.GetString(PartsOf(saved)["word/document.xml"]);
        Assert.Contains("commentRangeStart", body, StringComparison.Ordinal);
    }

    [Fact]
    public void EditingOneParagraphKeepsTrackedChangesOnTheOthers()
    {
        var original = Fixtures.WithTrackedChanges();
        var model = Clone(Open(original));

        Assert.True(EditFirstTextContaining(model, "intocado", "Mexido."));

        var (saved, _) = Save(original, model);
        var body = System.Text.Encoding.UTF8.GetString(PartsOf(saved)["word/document.xml"]);

        Assert.Contains("<w:ins ", body, StringComparison.Ordinal);
        Assert.Contains("<w:del ", body, StringComparison.Ordinal);
    }

    [Fact]
    public void SavingWithoutEditingRewritesNothing()
    {
        // Abrir e salvar sem mexer é o caso mais comum de todos — clicar em
        // salvar por reflexo. Precisa custar zero.
        var original = Fixtures.WithComment();
        var (_, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.True(result.PreservedBlocks > 0);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void OnlyTheDocumentBodyPartChanges()
    {
        // A promessa em uma linha: nada fora de word/document.xml é tocado.
        var original = Fixtures.WithAnchoredImage();
        var model = Clone(Open(original));
        EditFirstTextContaining(model, "Antes", "Texto novo.");

        var (saved, _) = Save(original, model);
        var before = PartsOf(original);
        var after = PartsOf(saved);

        var changed = before.Keys.Intersect(after.Keys)
            .Where(name => !before[name].SequenceEqual(after[name]))
            .ToList();

        Assert.Equal((string[])["word/document.xml"], changed);
        Assert.Empty(before.Keys.Except(after.Keys));
    }

    [Fact]
    public void EditingAParagraphThatCarriesACommentIsReportedAsLoss()
    {
        // Aqui a perda é real, e o aviso precisa dizer isso — com precisão, e
        // não como alerta genérico na abertura.
        var original = Fixtures.WithComment();
        var model = Clone(Open(original));

        Assert.True(EditFirstTextContaining(model, "Parágrafo comentado", "Reescrito."));

        var (_, result) = Save(original, model);

        Assert.Contains(result.Inventory.Lost, message => message.Contains("comentário", StringComparison.Ordinal));
    }

    // --- extração ------------------------------------------------------------

    [Fact]
    public void ReadsHeadingsAlignmentAndText()
    {
        var model = Open(Fixtures.Simple());

        Assert.Contains(Walk(model.Doc), n => n.Type == "heading");
        Assert.Contains("Segundo parágrafo", TextOf(model), StringComparison.Ordinal);
    }

    [Fact]
    public void ReadsAnchoredImageAsABlock()
    {
        // Toda imagem do corpus é âncora centralizada com a largura do texto —
        // o jeito do LibreOffice dizer "imagem no próprio parágrafo".
        var model = Open(Fixtures.WithAnchoredImage());

        var image = Walk(model.Doc).SingleOrDefault(n => n.Type == "image");
        Assert.NotNull(image);
        Assert.StartsWith("data:image/png;base64,", image!.Attrs!["src"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void ReadsSmallCapsAndCaps()
    {
        var model = Open(Fixtures.WithSmallCaps());
        var marks = Walk(model.Doc).SelectMany(n => n.Marks ?? []).Select(m => m.Type).ToList();

        Assert.Contains("smallCaps", marks);
        Assert.Contains("caps", marks);
    }

    [Fact]
    public void ReadsBulletListAsAList()
    {
        var model = Open(Fixtures.WithBulletList());

        Assert.Contains(Walk(model.Doc), n => n.Type == "bulletList");
        Assert.Equal(2, Walk(model.Doc).Count(n => n.Type == "listItem"));
    }

    [Fact]
    public void ReadsTable()
    {
        var model = Open(Fixtures.WithTable());

        Assert.Contains(Walk(model.Doc), n => n.Type == "table");
        Assert.Equal(4, Walk(model.Doc).Count(n => n.Type == "tableCell"));
    }

    [Fact]
    public void FlattensIdenticalSectionsWithoutReportingLoss()
    {
        // Sete seções idênticas são artefato do LibreOffice, não intenção do
        // autor. Avisar sobre elas encheria a tela de ruído em documento normal.
        var result = DocxReader.Read(Fixtures.WithIdenticalSections());

        Assert.Equal("A4", result.Model.Page.Size);
        Assert.DoesNotContain(result.Inventory.Lost, m => m.Contains("seções", StringComparison.Ordinal));
    }

    [Fact]
    public void ReportsLossWhenSectionsActuallyDiverge()
    {
        // A prova negativa do teste acima: sem ela, o achatamento poderia estar
        // engolindo divergência de verdade e ninguém saberia.
        var result = DocxReader.Read(Fixtures.WithDivergentSections());

        Assert.Contains(result.Inventory.Lost, m => m.Contains("seções", StringComparison.Ordinal));
    }

    [Fact]
    public void ReadsPageGeometryInMillimeters()
    {
        var page = Open(Fixtures.Simple()).Page;

        Assert.Equal("A4", page.Size);
        Assert.Equal("portrait", page.Orientation);
        Assert.Equal(25.4, page.Margins.Top, 1);
    }

    // --- estilos ------------------------------------------------------------

    private static Node FirstOfType(DocumentModelDto model, string type) =>
        Walk(model.Doc).First(n => n.Type == type);

    private static List<Mark> MarksOf(Node node) =>
        Walk(node).Where(n => n.Type == "text").SelectMany(n => n.Marks ?? []).ToList();

    /// <summary>
    /// O primeiro bloco que contém o texto dado.
    /// </summary>
    /// <remarks>
    /// Filtra por tipo de bloco de propósito: `Walk` inclui a raiz, e sem o
    /// filtro o `doc` inteiro casaria com qualquer busca — o teste passaria a
    /// medir o documento todo em vez do parágrafo.
    /// </remarks>
    private static Node BlockContaining(DocumentModelDto model, string needle) =>
        Walk(model.Doc)
            .Where(n => n.Type is "paragraph" or "heading")
            .First(n => Walk(n).Any(c => c.Text?.Contains(needle, StringComparison.Ordinal) == true));

    [Fact]
    public void ResolvesFormattingThatLivesOnlyInTheStyle()
    {
        // O caso que doeu na tela: `Heading1` no corpus real não é um título
        // grande, é uma barra vermelha com texto branco. Nada disso está no
        // parágrafo — ler só a formatação direta perdia tudo.
        var model = Open(Fixtures.WithStyles());
        var banner = FirstOfType(model, "paragraph");

        Assert.Equal("#943634", banner.Attrs!["background"]!.GetValue<string>());
        Assert.Equal("center", banner.Attrs["textAlign"]!.GetValue<string>());

        var style = MarksOf(banner).Single(m => m.Type == "textStyle");
        Assert.Equal("#ffffff", style.Attrs!["color"]!.GetValue<string>());
        Assert.Equal("Arial", style.Attrs["fontFamily"]!.GetValue<string>());
        Assert.Equal("10pt", style.Attrs["fontSize"]!.GetValue<string>());
        Assert.Contains(MarksOf(banner), m => m.Type == "bold");
    }

    [Fact]
    public void FollowsTheBasedOnChain()
    {
        // `Corpo` herda a fonte de `Base` e só acrescenta o alinhamento.
        var model = Open(Fixtures.WithStyles());
        var body = BlockContaining(model, "Texto do corpo");

        Assert.Equal("justify", body.Attrs!["textAlign"]!.GetValue<string>());
        Assert.Equal("Arial", MarksOf(body).Single(m => m.Type == "textStyle").Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void DirectFormattingBeatsTheStyle()
    {
        var model = Open(Fixtures.WithStyles());
        var overridden = BlockContaining(model, "à direita");

        Assert.Equal("right", overridden.Attrs!["textAlign"]!.GetValue<string>());
    }

    [Fact]
    public void ParagraphMarkFormattingDoesNotReachTheRuns()
    {
        // `w:rPr` dentro de `w:pPr` formata a marca de parágrafo — o "¶" — e
        // não o texto. Aplicá-la aos runs punha negrito em parágrafos inteiros
        // que no Word aparecem normais.
        var model = Open(Fixtures.WithStyles());
        var paragraph = BlockContaining(model, "negrito");

        Assert.DoesNotContain(MarksOf(paragraph), m => m.Type == "bold");
    }

    [Fact]
    public void PutsTheFontOnTheBlockAndNotOnlyOnTheRuns()
    {
        // A altura da linha nasce da fonte **do elemento**, não do texto dentro
        // dele. Sem a fonte no bloco, um parágrafo de 10 pt continuava ocupando
        // os 12 pt que a folha do editor declara — e um título de 10 pt virava
        // uma barra alta demais, porque o editor desenha títulos em 22 pt.
        var model = Open(Fixtures.WithStyles());
        var banner = FirstOfType(model, "paragraph");

        Assert.Equal("10pt", banner.Attrs!["fontSize"]!.GetValue<string>());
        Assert.Equal("Arial", banner.Attrs["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsKeepWithNext()
    {
        // Diz que o bloco não fica sozinho no pé da página. A marca de fim de
        // página e a folha de impressão usam o mesmo sinal — se divergissem, a
        // marca cairia num lugar e o PDF quebraria noutro.
        var model = Open(Fixtures.WithKeepNext());
        var kept = BlockContaining(model, "Rótulo");

        Assert.True(kept.Attrs!["keepNext"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Solto").Attrs?.ContainsKey("keepNext") ?? false);
    }

    [Fact]
    public void SurvivesAStyleThatInheritsFromItself()
    {
        // Documento é dado não confiável: uma cadeia circular não pode travar.
        var model = Open(Fixtures.WithCircularStyle());

        Assert.Contains("Texto.", TextOf(model), StringComparison.Ordinal);
    }

    // --- inventário ---------------------------------------------------------

    [Fact]
    public void ReportsCommentsAsInvisibleNotLost()
    {
        // A distinção é o ponto: comentários são preservados no arquivo, e o
        // editor não os mostra. Chamar isso de perda seria mentira, e o usuário
        // aprenderia a ignorar o aviso.
        var result = DocxReader.Read(Fixtures.WithComment());

        Assert.Contains("comentários", result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void ReportsTrackedChangesAsInvisible()
    {
        var result = DocxReader.Read(Fixtures.WithTrackedChanges());

        Assert.Contains("controle de alterações", result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Lost);
    }

    // --- arquivos problemáticos ---------------------------------------------

    [Theory]
    [InlineData("não é um zip")]
    [InlineData("PK zip truncado")]
    public void RejectsGarbageWithAReadableMessage(string garbage)
    {
        // Documento é dado não confiável (spec): qualquer entrada vira frase,
        // nunca processo derrubado.
        var problem = Assert.Throws<DocxException>(
            () => DocxReader.Read(System.Text.Encoding.UTF8.GetBytes(garbage)));

        Assert.DoesNotContain("Exception", problem.Message, StringComparison.Ordinal);
        Assert.Contains("Word", problem.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsAZipThatIsNotADocument()
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            archive.CreateEntry("qualquer/coisa.txt");
        }

        Assert.Throws<DocxException>(() => DocxReader.Read(buffer.ToArray()));
    }
}

using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Nodes;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// A promessa da gravação cirúrgica: editar um documento não pode custar o que
/// não foi editado.
/// </summary>
public class DocxRoundTripTests
{
    // Os passos do caminho moram em Roundtrip, comuns a DocxWriteBackTests.
    private static DocumentModelDto Open(byte[] bytes) => Roundtrip.Open(bytes);

    private static DocumentModelDto OpenFlat(byte[] bytes) => Roundtrip.OpenFlat(bytes);

    private static (byte[] Bytes, SaveResult Result) Save(byte[] original, DocumentModelDto model) =>
        Roundtrip.Save(original, model);

    private static DocumentModelDto Clone(DocumentModelDto model) => Roundtrip.Clone(model);

    private static IEnumerable<Node> Walk(Node node) => Roundtrip.Walk(node);

    private static string TextOf(DocumentModelDto model) => Roundtrip.TextOf(model);

    private static bool EditFirstTextContaining(DocumentModelDto model, string needle, string replacement) =>
        Roundtrip.EditFirstTextContaining(model, needle, replacement);

    private static Dictionary<string, byte[]> PartsOf(byte[] docx) => Roundtrip.PartsOf(docx);

    [Fact]
    public void EditingOneParagraphKeepsCommentsOnTheOthers()
    {
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
        // Abrir e salvar sem mexer, o caso mais comum, custa zero.
        var original = Fixtures.WithComment();
        var (_, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.True(result.PreservedBlocks > 0);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void OnlyTheDocumentBodyPartChanges()
    {
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
    public void EditingAParagraphThatCarriesACommentKeepsTheAnchor()
    {
        // A âncora volta com o parágrafo; ver CommentsTests.
        var original = Fixtures.WithComment();
        var model = Clone(Open(original));

        Assert.True(EditFirstTextContaining(model, "Parágrafo comentado", "Reescrito."));

        var (saved, result) = Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        var body = System.Text.Encoding.UTF8.GetString(PartsOf(saved)["word/document.xml"]);
        Assert.Contains("<w:commentRangeStart w:id=\"1\"", body, StringComparison.Ordinal);
        Assert.Contains("<w:commentReference w:id=\"1\"", body, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadsHeadingsAlignmentAndText()
    {
        var model = Open(Fixtures.Simple());

        Assert.Contains(Walk(model.Doc), n => n.Type == "heading");
        Assert.Contains("Segundo parágrafo", TextOf(model), StringComparison.Ordinal);
    }

    [Fact]
    public void ImagemAncoradaSaiDoFluxo()
    {
        // No Word o objeto ancorado mora numa posição da folha e não empurra o texto.
        var model = Open(Fixtures.WithAnchoredImage());

        Assert.DoesNotContain(Walk(model.Doc), n => n.Type == "image");

        var floats = FloatsOf(model.Doc.Content![1]);
        var image = Assert.Single(floats);
        Assert.Equal("image", image.GetProperty("kind").GetString());
        Assert.StartsWith(
            "data:image/png;base64,",
            image.GetProperty("src").GetString()!,
            StringComparison.Ordinal);
    }

    [Fact]
    public void AncoradaOndeOFluxoAPoriaContinuaNoFluxo()
    {
        // O LibreOffice grava a imagem no próprio parágrafo como `wp:anchor` sem
        // deslocamento, centralizada: ela ocupa altura no fluxo.
        var model = Open(Fixtures.WithAnchoredImageInTheFlow());

        var image = Assert.Single(Walk(model.Doc).Where(n => n.Type == "image"));
        Assert.Equal(554, image.Attrs!["width"]!.GetValue<int>());

        Assert.All(model.Doc.Content!, block => Assert.Empty(FloatsOf(block)));
    }

    [Fact]
    public void ImagemNoFluxoContinuaSendoBloco()
    {
        // `wp:inline` ocupa lugar na linha.
        var model = Open(Fixtures.WithInlineImage());

        var image = Assert.Single(Walk(model.Doc).Where(n => n.Type == "image"));
        Assert.Equal(554, image.Attrs!["width"]!.GetValue<int>());

        // Sem a altura o navegador reserva zero até decodificar, e a paginação mede antes.
        Assert.Equal(277, image.Attrs["height"]!.GetValue<int>());
    }

    [Fact]
    public void ImagemReescritaMantemOTamanhoQueTinha()
    {
        // O parágrafo editado é regravado do modelo, e a imagem guarda a forma.
        var original = Fixtures.WithInlineImage();
        var model = Clone(Open(original));

        var paragraph = Walk(model.Doc).First(n => n.Type == "paragraph");
        paragraph.Attrs!.Remove("oid");

        var (bytes, _) = Save(original, model);
        var image = Assert.Single(Walk(Open(bytes).Doc).Where(n => n.Type == "image"));

        Assert.Equal(554, image.Attrs!["width"]!.GetValue<int>());
        Assert.Equal(277, image.Attrs["height"]!.GetValue<int>());
    }

    [Fact]
    public void AListaHerdaOEspacamentoDosParagrafosDela()
    {
        // No arquivo a lista são parágrafos numerados, cada um com seu espaçamento;
        // o elemento da lista no editor não pode somar o espaçamento padrão.
        var model = Open(Fixtures.WithBulletList());
        var list = Walk(model.Doc).First(n => n.Type == "bulletList");

        Assert.Equal(0, list.Attrs!["spaceBefore"]!.GetValue<double>());
        Assert.Equal(0, list.Attrs["spaceAfter"]!.GetValue<double>());
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
    public void ReadsSuperscriptAndSubscript()
    {
        var model = Open(Fixtures.WithVerticalAlignment());
        var marks = Walk(model.Doc).SelectMany(n => n.Marks ?? []).Select(m => m.Type).ToList();

        Assert.Contains("superscript", marks);
        Assert.Contains("subscript", marks);
    }

    [Fact]
    public void SavingVerticalAlignmentWithoutEditingRewritesNothing()
    {
        // Ler um recurso novo não pode fazer o modelo divergir do arquivo, senão
        // abrir e salvar reescreve o parágrafo.
        var original = Fixtures.WithVerticalAlignment();
        var (_, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.True(result.PreservedBlocks > 0);
        Assert.Empty(result.Inventory.Lost);
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
    public void ReadsColumnWidthsShadingAndHeaderRow()
    {
        // A largura de coluna vem do `w:tblGrid`, que o TableKit redimensiona.
        var table = Walk(Open(Fixtures.WithStyledTable()).Doc).First(n => n.Type == "table");
        var header = table.Content![0].Content![0];

        // `w:tblHeader` é a linha que o Word repete no alto de cada página.
        Assert.Equal("tableHeader", header.Type);
        Assert.Equal("tableCell", table.Content[1].Content![0].Type);

        // 15 twips por pixel: 4000 são 267 px.
        Assert.Equal("[267]", header.Attrs!["colwidth"]!.ToJsonString());
        Assert.Equal("[333]", table.Content[0].Content![1].Attrs!["colwidth"]!.ToJsonString());
        Assert.Equal("#d9d9d9", header.Attrs["shading"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsCellBordersInTheCanonicalForm()
    {
        // Uma escrita só, comparada com a do editor: lado, estilo, espessura em pontos e cor.
        var model = Open(Fixtures.WithPatternedCell());
        var cell = Walk(model.Doc).First(n => n.Type == "tableCell");

        // 24 oitavos de ponto são 3 pt; `thickThinSmallGap` vira linha simples, e a trama, cor lisa.
        Assert.Equal("top:single,3,#000000", cell.Attrs!["borders"]!.GetValue<string>());
        Assert.Equal("#ffff00", cell.Attrs["shading"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsImageAlternativeText()
    {
        // `wp:docPr/@descr` é o que o leitor de tela lê no lugar da imagem.
        var image = Walk(Open(Fixtures.WithDescribedImage()).Doc).First(n => n.Type == "image");

        Assert.Equal("Organograma da diretoria", image.Attrs!["alt"]!.GetValue<string>());
    }

    [Fact]
    public void FlattensIdenticalSectionsWithoutReportingLoss()
    {
        // Sete seções idênticas são artefato do LibreOffice, e não pedem aviso.
        var result = DocxReader.Read(Fixtures.WithIdenticalSections());

        Assert.Equal("A4", result.Model.Page.Size);
        Assert.DoesNotContain(result.Inventory.Lost, m => m.Contains("seções", StringComparison.Ordinal));
    }

    [Fact]
    public void ReadsPageGeometryInMillimeters()
    {
        var page = Open(Fixtures.Simple()).Page;

        Assert.Equal("A4", page.Size);
        Assert.Equal("portrait", page.Orientation);
        Assert.Equal(25.4, page.Margins.Top, 1);
    }

    /// <summary>`word/document.xml` de dentro do pacote.</summary>
    private static string MainDocumentXml(byte[] bytes)
    {
        using var buffer = new MemoryStream(bytes, writable: false);
        using var package = new System.IO.Compression.ZipArchive(buffer);
        using var stream = package.GetEntry("word/document.xml")!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>Os objetos ancorados de um bloco.</summary>
    private static List<JsonElement> FloatsOf(Node node)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue("floats", out var value) || value is null)
        {
            return [];
        }

        return [.. JsonDocument.Parse(value.ToJsonString()).RootElement.EnumerateArray()];
    }

    /// <summary>Uma propriedade de texto do objeto, ou nulo quando ausente.</summary>
    private static string? TextOf(JsonElement item, string name) =>
        item.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    /// <summary>Todo o texto de dentro de uma caixa.</summary>
    private static string TextOfFloat(JsonElement item) =>
        item.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array
            ? string.Concat(content.EnumerateArray().Select(TextOfJson))
            : string.Empty;

    private static string TextOfJson(JsonElement node)
    {
        if (node.TryGetProperty("text", out var text)) return text.GetString() ?? string.Empty;
        return node.TryGetProperty("content", out var children) && children.ValueKind == JsonValueKind.Array
            ? string.Concat(children.EnumerateArray().Select(TextOfJson))
            : string.Empty;
    }

    /// <summary>Todo o texto de uma faixa, nas três colunas.</summary>
    private static string TextOfBand(BandDto? band) =>
        band is null
            ? string.Empty
            : string.Concat(
                band.Left.Concat(band.Center).Concat(band.Right).Select(piece => piece.Text ?? string.Empty));

    private static Node FirstOfType(DocumentModelDto model, string type) =>
        Walk(model.Doc).First(n => n.Type == type);

    private static List<Mark> MarksOf(Node node) =>
        Walk(node).Where(n => n.Type == "text").SelectMany(n => n.Marks ?? []).ToList();

    /// <summary>
    /// O primeiro bloco com o texto. Filtra por tipo porque <c>Walk</c> inclui a raiz,
    /// que casaria com qualquer busca.
    /// </summary>
    private static Node BlockContaining(DocumentModelDto model, string needle) =>
        Walk(model.Doc)
            .Where(n => n.Type is "paragraph" or "heading")
            .First(n => Walk(n).Any(c => c.Text?.Contains(needle, StringComparison.Ordinal) == true));

    [Fact]
    public void ResolvesFormattingThatLivesOnlyInTheStyle()
    {
        // O `Heading1` do corpus é uma barra vermelha de texto branco, toda no estilo.
        var model = OpenFlat(Fixtures.WithStyles());
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
        var model = OpenFlat(Fixtures.WithStyles());
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
        // `w:rPr` em `w:pPr` formata a marca de parágrafo, e não o texto.
        var model = Open(Fixtures.WithStyles());
        var paragraph = BlockContaining(model, "negrito");

        Assert.DoesNotContain(MarksOf(paragraph), m => m.Type == "bold");
    }

    [Fact]
    public void PutsTheFontOnTheBlockAndNotOnlyOnTheRuns()
    {
        // A altura da linha nasce da fonte do elemento, e não do texto dentro dele.
        var model = OpenFlat(Fixtures.WithStyles());
        var banner = FirstOfType(model, "paragraph");

        Assert.Equal("10pt", banner.Attrs!["fontSize"]!.GetValue<string>());
        Assert.Equal("Arial", banner.Attrs["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void ParagraphWithoutAStyleStillGetsTheDefaultOne()
    {
        // O estilo `w:default="1"` vale para quem não declara `w:pStyle`.
        var model = OpenFlat(Fixtures.WithLineMetrics());
        var first = BlockContaining(model, "estilo padrão");

        Assert.Equal("12pt", first.Attrs!["fontSize"]!.GetValue<string>());
    }

    [Fact]
    public void TheParagraphMarkGivesTheBlockItsFont()
    {
        // É com a marca de parágrafo que o Word mede a linha, também a do parágrafo vazio.
        var model = Open(Fixtures.WithLineMetrics());
        var marked = BlockContaining(model, "Verdana");

        Assert.Equal("Verdana", marked.Attrs!["fontFamily"]!.GetValue<string>());
        Assert.Equal("10pt", marked.Attrs["fontSize"]!.GetValue<string>());

        // E ela continua fora dos runs.
        Assert.DoesNotContain(MarksOf(marked), m => m.Type == "textStyle");
        var flat = BlockContaining(OpenFlat(Fixtures.WithLineMetrics()), "Verdana");
        var run = MarksOf(flat).Single(m => m.Type == "textStyle");
        Assert.Equal("Times New Roman", run.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void SilenceAboutSpacingMeansSingleAndZero()
    {
        // Arquivo calado diz zero de espaço e espaçamento simples, e não o padrão do editor.
        var model = OpenFlat(Fixtures.WithLineMetrics());
        var first = BlockContaining(model, "estilo padrão");

        // Simples sai como número, a altura natural da fonte: o Chromium arredonda o
        // `normal` para pixel inteiro.
        Assert.Equal("1.1499", first.Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal(0, first.Attrs["spaceBefore"]!.GetValue<double>());
        Assert.Equal(0, first.Attrs["spaceAfter"]!.GetValue<double>());
    }

    [Fact]
    public void ReadsLineSpacingThatIsLockedInPoints()
    {
        // `exact` e `atLeast` dizem a altura em twips.
        var model = Open(Fixtures.WithLineMetrics());

        Assert.Equal("9pt", BlockContaining(model, "travada").Attrs!["lineHeight"]!.GetValue<string>());

        // O múltiplo é sobre a altura natural: uma vez e meia de Liberation Serif é 1,7248 em.
        Assert.Equal("1.7248", BlockContaining(model, "uma vez e meia").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void OMultiploDaEntrelinhaEVezAAlturaDaFonte()
    {
        // "1,13 linha" é 1,13 vez a linha simples da fonte, que em Arial é 1,1499 em.
        var model = Open(Fixtures.WithLineMetrics());

        // 271/240 a quatro casas (1,1292), a grade do arquivo, sobre 1,1499.
        Assert.Equal("1.2985", BlockContaining(model, "Arial e um pouco mais").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void FonteQueNaoVaiNoInstaladorAindaRespeitaOMultiplo()
    {
        // Fonte desconhecida com múltiplo declarado: o palpite de 1,15, a altura de
        // quase toda fonte latina, porque sobre o tamanho da fonte erra 15 %.
        var model = Open(Fixtures.WithLineMetrics());

        // 271/240 sobre o palpite de 1,15.
        Assert.Equal("1.2985", BlockContaining(model, "fonte que ninguém tem").Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal("normal", BlockContaining(model, "Verdana de dez").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void AFonteAusenteCaiNaSubstitutaDoTipoCerto()
    {
        // `word/fontTable.xml` diz o tipo da fonte que pode faltar, para a pilha do CSS.
        var model = Open(Fixtures.WithMissingFont());

        var titulo = MarksOf(BlockContaining(model, "Título da capa")).Single(m => m.Type == "textStyle");
        Assert.Equal("Segoe UI, sans-serif", titulo.Attrs!["fontFamily"]!.GetValue<string>());

        // Sem tipo declarado, sem pilha inventada.
        var outro = MarksOf(BlockContaining(model, "Sem tipo")).Single(m => m.Type == "textStyle");
        Assert.Equal("Fonte Fantasma", outro.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void SoONomeDaFonteVoltaParaOArquivo()
    {
        // `w:rFonts` guarda o nome, e não a pilha: senão a fonte seria "Segoe UI, sans-serif".
        var original = Fixtures.WithMissingFont();
        var model = Clone(Open(original));

        var paragraph = Walk(model.Doc).First(n => n.Type == "paragraph");
        paragraph.Attrs!.Remove("oid");

        var (bytes, _) = Save(original, model);
        var titulo = MarksOf(BlockContaining(Open(bytes), "Título da capa")).Single(m => m.Type == "textStyle");

        Assert.Equal("Segoe UI, sans-serif", titulo.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsKeepWithNext()
    {
        // A marca de fim de página e a impressão usam o mesmo sinal, senão cortariam em lugares diferentes.
        var model = Open(Fixtures.WithKeepNext());
        var kept = BlockContaining(model, "Rótulo");

        Assert.True(kept.Attrs!["keepNext"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Solto").Attrs?.ContainsKey("keepNext") ?? false);
    }

    [Fact]
    public void PutsTheTopAnchoredImageBeforeTheText()
    {
        // O quadro ancorado ao topo vem antes do texto, como o Word e o LibreOffice o desenham.
        var model = Open(Fixtures.WithTextAroundTopAnchoredImage());
        var paragraph = model.Doc.Content!.Single(node => node.Type == "paragraph");

        Assert.Equal(["image", "text", "text"], paragraph.Content!.Select(node => node.Type));
    }

    [Fact]
    public void ResolvesCellMarginsLikeWord()
    {
        // Lado a lado: a tabela, o estilo padrão de tabela, e o 0/108 do Word.
        var model = Open(Fixtures.WithCellMargins());
        var tables = model.Doc.Content!.Where(node => node.Type == "table").ToList();

        Assert.Equal("100 108 60 108", tables[0].Attrs!["cellMargins"]!.GetValue<string>());
        Assert.Equal("0 108 60 108", tables[1].Attrs!["cellMargins"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsKeepLines()
    {
        var model = Open(Fixtures.WithKeepLines());

        Assert.True(BlockContaining(model, "Linhas juntas").Attrs!["keepLines"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Comum").Attrs?.ContainsKey("keepLines") ?? false);
    }

    [Fact]
    public void ReadsWidowControlOnlyWhenTurnedOff()
    {
        // Ligado é o padrão do Word: só o que o desliga diz alguma coisa.
        var model = Open(Fixtures.WithKeepLines());

        Assert.False(BlockContaining(model, "Viúva permitida").Attrs!["widowControl"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Comum").Attrs?.ContainsKey("widowControl") ?? false);
    }

    [Fact]
    public void TreatsLeadingTabsAsCentering()
    {
        // No corpus o título vem alinhado à esquerda com tabulações até uma parada
        // centralizada; no HTML a tabulação colapsa.
        var model = Open(Fixtures.WithTabCentering());
        var centered = BlockContaining(model, "Centralizado por tabulação");

        Assert.Equal("center", centered.Attrs!["textAlign"]!.GetValue<string>());

        // As tabulações deste parágrafo viram posicionamento; a do meio da linha fica.
        var text = string.Concat(Walk(centered).Where(n => n.Type == "text").Select(n => n.Text));
        Assert.Equal("Centralizado por tabulação", text);
    }

    [Fact]
    public void LeavesTabsInTheMiddleOfALineAlone()
    {
        // Só a tabulação do começo da linha: "esquerda [tab] meio" é outra coisa.
        var model = Open(Fixtures.WithTabCentering());
        var inline = BlockContaining(model, "Esquerda");

        Assert.NotEqual("center", inline.Attrs?["textAlign"]?.GetValue<string>());
    }

    [Fact]
    public void SurvivesAStyleThatInheritsFromItself()
    {
        // Documento é dado não confiável: cadeia circular não trava.
        var model = Open(Fixtures.WithCircularStyle());

        Assert.Contains("Texto.", TextOf(model), StringComparison.Ordinal);
    }

    [Fact]
    public void CommentsAreNeitherInvisibleNorLost()
    {
        // O painel os mostra: nem aviso de invisível, nem de perda.
        var result = DocxReader.Read(Fixtures.WithComment());

        Assert.DoesNotContain("comentários", result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void TrackedChangesNoLongerLockOrWarn()
    {
        // A revisão de texto é marca do editor: nem invisível, nem trava.
        var result = DocxReader.Read(Fixtures.WithTrackedChanges());

        Assert.Empty(result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Structural);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void RevisaoDeEstruturaEEstruturalEComentarioNao()
    {
        // A lista que decide o somente leitura é subconjunto da invisibilidade:
        // comentário e revisão de texto não entram, a revisão de estrutura entra.
        var comentado = DocxReader.Read(Fixtures.WithComment()).Inventory;
        var revisado = DocxReader.Read(Fixtures.WithInsertedCell()).Inventory;

        Assert.DoesNotContain(Inventory.Comments, comentado.Structural);
        Assert.Contains(Inventory.StructureRevisions, revisado.Structural);
        // Tudo que é estrutural é invisível, senão o aviso deixaria de mencioná-lo.
        Assert.All(revisado.Structural, item => Assert.Contains(item, revisado.Invisible));
    }

    [Fact]
    public void DocumentoComumNaoAbreEmSomenteLeitura()
    {
        // Somente leitura em documento normal ensinaria a clicar "editar mesmo assim" sem ler.
        var result = DocxReader.Read(Fixtures.Simple());

        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void ImagemAncoradaNaoEEstrutural()
    {
        // Perda de aparência, e não de conteúdo: quase todo documento do corpus tem uma.
        var result = DocxReader.Read(Fixtures.WithAnchoredImage());

        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void DesenhoAncoradoNoCabecalhoNaoEspremeEmTresColunas()
    {
        // Desenho ancorado na faixa vai como objeto, com posição e giro: a marca
        // lateral do corpus fica em pé, 28,6 mm numa banda de 10 mm.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(4563177)).Model.Page;

        Assert.Empty(page.Header!.Left);
        Assert.Empty(page.Header.Center);
        Assert.Empty(page.Header.Right);

        var logo = Assert.Single(page.Header.Floats!);
        Assert.Equal("column", logo.HorizontalFrom);
        Assert.Equal(126.75, logo.HorizontalOffsetMm!.Value, 2);
    }

    [Fact]
    public void ODeslocamentoNegativoDoCabecalhoSobreviveComSinal()
    {
        // Marca que sai da coluna pela esquerda: o sinal conta.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(-1123950)).Model.Page;

        var marca = Assert.Single(page.Header!.Floats!);
        Assert.Equal(-31.22, marca.HorizontalOffsetMm!.Value, 2);
    }

    [Fact]
    public void CabecalhoEmGradeSaiComoGrade()
    {
        // O cabeçalho corporativo do corpus é uma tabela, com o logotipo mesclado em
        // três linhas: três colunas não bastam.
        var page = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page;
        var rows = page.Header!.Rows!;

        Assert.Equal(3, rows.Count);
        Assert.Equal(
            ["Chamado 10001", "Data de revisão", "Título do documento", "30/07/2026",
             "Rodapé do cabeçalho", "Página", "Revisão"],
            rows.SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Pieces)
                .Where(piece => piece.Kind == PieceDto.KindText)
                .Select(piece => piece.Text));
    }

    [Fact]
    public void ACelulaMescladaCresceEmVezDeDeixarBuracos()
    {
        // A mesclagem vertical do OOXML é `restart` em cima e células vazias embaixo.
        var rows = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page.Header!.Rows!;

        Assert.Equal(3, rows[0].Cells[0].RowSpan);
        Assert.Equal("image", rows[0].Cells[0].Pieces[0].Kind);

        Assert.Equal(2, rows[1].Cells.Count);
        Assert.Equal(3, rows[2].Cells.Count);
    }

    [Fact]
    public void AGradeTrazLarguraJuntoEBordaResolvida()
    {
        // Largura em fração da grade, porque quem desenha não sabe quanto vale um twip;
        // a borda vem resolvida, porque cada lado sai de três lugares no OOXML.
        var rows = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page.Header!.Rows!;

        Assert.Equal(0.2, rows[0].Cells[0].Width, 3);
        Assert.Equal(2, rows[0].Cells[2].Span);

        Assert.Contains("b", rows[0].Cells[1].Borders, StringComparison.Ordinal);
        // `w:nil` na célula apaga o risco que a tabela pediu.
        Assert.DoesNotContain("b", rows[2].Cells[1].Borders, StringComparison.Ordinal);
    }

    [Fact]
    public void OGrupoDoCabecalhoSaiPecaPorPeca()
    {
        // A âncora dá o lugar do grupo, e `a:chOff`/`a:chExt` a régua de dentro.
        var page = DocxReader.Read(Fixtures.WithHeaderGroup()).Model.Page;
        var floats = page.Header!.Floats!;

        // Três peças: o logotipo, a caixa do título e o filete de baixo.
        Assert.Equal(3, floats.Count);

        var logo = floats.Single(item => item.Kind == "image");
        Assert.Equal(48, logo.WidthMm, 1);
        Assert.Equal(10.5, logo.HeightMm, 1);
        Assert.Equal(129, logo.DxMm, 1);
        Assert.Equal(0, logo.DyMm, 1);

        var titulo = floats.Single(item => item.Kind == "text");
        Assert.Equal(84.8, titulo.WidthMm, 1);
        Assert.Equal(41.7, titulo.DxMm, 1);
        Assert.Equal(6.9, titulo.DyMm, 1);
        Assert.Equal(
            "EVIDÊNCIAS DO ROTEIRO",
            titulo.Content!.Single().Content!.Single().Text);
    }

    [Fact]
    public void ADistanciaDaFaixaAteABordaEhLida()
    {
        // A origem vertical das âncoras da faixa, que se dizem relativas ao parágrafo.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(0)).Model.Page;

        Assert.Equal(12.5, page.HeaderDistanceMm, 1);
    }

    [Fact]
    public void QuebraDentroDoParagrafoViraPropriedadeDoBloco()
    {
        // Nó de bloco dentro da linha o editor não aceita, e `<div>` em `<p>` o HTML
        // desaloja: os índices deixariam de casar com o papel.
        var model = Open(Fixtures.WithBreakInsideParagraph());

        var blocos = model.Doc.Content!;
        Assert.DoesNotContain(Walk(model.Doc), n => n.Type == "pageBreak");
        Assert.True(blocos[0].Attrs!["breakAfter"]!.GetValue<bool>());
    }

    [Fact]
    public void AQuebraDoParagrafoVoltaAoArquivoAoEditar()
    {
        // O editado precisa escrever a quebra de volta; o intocado a traz no XML original.
        var bytes = Fixtures.WithBreakInsideParagraph();
        var model = Clone(Open(bytes));
        Assert.True(EditFirstTextContaining(model, "Fim da", "Outro texto."));

        var reaberto = Open(Save(bytes, model).Bytes);

        Assert.Contains("Outro texto.", TextOf(reaberto), StringComparison.Ordinal);
        Assert.True(reaberto.Doc.Content![0].Attrs!["breakAfter"]!.GetValue<bool>());
    }

    [Fact]
    public void OCabecalhoPadraoNaoDependeDaOrdemDeGravacao()
    {
        // O que vale é o `w:type`, e não a ordem das referências no XML.
        var bytes = Fixtures.WithFirstPageHeader(titlePage: true);

        // A armadilha: a referência `first`, não vazia, vem antes da `default`.
        var xml = MainDocumentXml(bytes);
        Assert.InRange(
            xml.IndexOf("w:type=\"first\"", StringComparison.Ordinal),
            0,
            xml.IndexOf("w:type=\"default\"", StringComparison.Ordinal));

        Assert.Equal("Miolo", TextOfBand(DocxReader.Read(bytes).Model.Page.Header));
    }

    [Fact]
    public void PrimeiraPaginaComCabecalhoProprioEhLida()
    {
        var page = DocxReader.Read(Fixtures.WithFirstPageHeader(titlePage: true)).Model.Page;

        Assert.Equal("Capa", TextOfBand(page.FirstHeader));
    }

    [Fact]
    public void SemTitlePgOCabecalhoDaCapaNaoEhUsado()
    {
        // O Word guarda a parte `first` com o interruptor desligado (quatro dos seis
        // documentos do corpus): ela vem, desligada, e quem desenha não a usa.
        var page = DocxReader.Read(Fixtures.WithFirstPageHeader(titlePage: false)).Model.Page;

        Assert.False(page.TitlePage);
        Assert.Equal("Capa", TextOfBand(page.FirstHeader));
        Assert.Equal("Miolo", TextOfBand(page.Header));
    }

    // O leitor e o editor descrevem o mesmo bloco de formas diferentes: cada teste
    // abaixo fixa uma delas como "não é edição".
    [Fact]
    public void AtributoNuloEAusenteSaoAMesmaCoisa()
    {
        // O ProseMirror materializa todo atributo do schema, com `null`.
        var doLeitor = Node.Of("paragraph").With("fontSize", "12pt");
        var doEditor = Node.Of("paragraph").With("fontSize", "12pt").With("styleId", null);

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void AOrdemDasChavesNaoContaComoEdicao()
    {
        // `JsonObject` guarda a ordem de inserção, que difere entre os dois lados.
        var doLeitor = Node.Of("paragraph").With("lineHeight", 1.16).With("fontSize", "12pt");
        var doEditor = Node.Of("paragraph").With("fontSize", "12pt").With("lineHeight", 1.16);

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void AOrdemDasMarcasNaoContaComoEdicao()
    {
        // O ProseMirror ordena as marcas pela posição no schema.
        var doLeitor = new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("bold"), Mark.Of("textStyle", "color", "#404040")] };
        var doEditor = new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("textStyle", "color", "#404040"), Mark.Of("bold")] };

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void TextoVizinhoComAsMesmasMarcasEUmTextoSo()
    {
        // O leitor emite um nó por `w:r`: "Acme® Software" chega partido.
        var doLeitor = Node.Of(
            "paragraph",
            new Node { Type = "text", Text = "Acme" },
            new Node { Type = "text", Text = "® Software" });
        var doEditor = Node.Of("paragraph", new Node { Type = "text", Text = "Acme® Software" });

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void TextoVizinhoComMarcasDiferentesNaoSeFunde()
    {
        // Fundir sem olhar as marcas esconderia uma edição de formatação.
        var negrito = Node.Of(
            "paragraph",
            new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("bold")] },
            new Node { Type = "text", Text = "® Software" });
        var liso = Node.Of("paragraph", new Node { Type = "text", Text = "Acme® Software" });

        Assert.NotEqual(negrito.Fingerprint(), liso.Fingerprint());
    }

    [Fact]
    public void EdicaoDeVerdadeContinuaSendoVista()
    {
        // Afrouxar demais preserva o XML de um bloco mudado, o que é perda de dados.
        var antes = Node.Of("paragraph", new Node { Type = "text", Text = "Sumário" });
        var depois = Node.Of("paragraph", new Node { Type = "text", Text = "Sumario" });

        Assert.NotEqual(antes.Fingerprint(), depois.Fingerprint());
    }

    [Fact]
    public void OidNaoEntraNaComparacao()
    {
        // Identidade não é conteúdo.
        var primeiro = Node.Of("paragraph", new Node { Type = "text", Text = "Texto" }).With("oid", "b1");
        var segundo = Node.Of("paragraph", new Node { Type = "text", Text = "Texto" }).With("oid", "b9");

        Assert.Equal(primeiro.Fingerprint(), segundo.Fingerprint());
    }

    [Fact]
    public void CaixaDeTextoViraObjetoComOTextoDentro()
    {
        // A caixa é um fluxo de texto próprio: na linha, título e subtítulo se emendariam.
        var floats = FloatsOf(Open(Fixtures.WithTextBoxes()).Doc.Content![0]);

        Assert.Equal(2, floats.Count);
        Assert.All(floats, item => Assert.Equal("text", item.GetProperty("kind").GetString()));
        Assert.Contains(floats, item => TextOfFloat(item) == "Título do manual");
        Assert.Contains(floats, item => TextOfFloat(item) == "Subtítulo do manual");
    }

    [Fact]
    public void CaixaDeTextoNaoEntraDuasVezes()
    {
        // DrawingML e VML de reserva: um ramo só, senão dois objetos no mesmo lugar.
        var floats = FloatsOf(Open(Fixtures.WithTextBoxes()).Doc.Content![0]);

        Assert.Single(floats.Where(item => TextOfFloat(item) == "Título do manual"));
    }

    [Fact]
    public void CaixaDeTextoSobreviveAEdicaoDoParagrafoAncora()
    {
        // O escritor não gera forma do zero: copia os objetos ancorados do XML original do bloco.
        var original = Fixtures.WithTextBoxes();
        var model = Clone(Open(original));

        // Editar o texto, e não apagar o `oid`, que é o fio até o XML original.
        var paragraph = Walk(model.Doc).First(n => n.Type == "paragraph");
        paragraph.Content = [new Node { Type = "text", Text = "Capa" }];

        var (bytes, _) = Save(original, model);
        var floats = FloatsOf(Open(bytes).Doc.Content![0]);

        Assert.Contains(floats, item => TextOfFloat(item) == "Título do manual");
        Assert.Contains(floats, item => TextOfFloat(item) == "Subtítulo do manual");
    }

    [Fact]
    public void TextoDigitadoNaCaixaVoltaParaOArquivo()
    {
        // O título e o subtítulo da capa moram em caixas: o texto novo tem de voltar.
        var original = Fixtures.WithTextBoxes();
        var model = Clone(Open(original));

        var capa = model.Doc.Content![0];
        var floats = FloatsOf(capa);
        var editado = JsonNode.Parse(floats[0].GetRawText())!.AsObject();
        editado["content"]![0]!["content"]![0]!["text"] = "Título trocado";
        capa.Attrs!["floats"] = new JsonArray(editado, JsonNode.Parse(floats[1].GetRawText()));

        var (bytes, _) = Save(original, model);
        var depois = FloatsOf(Open(bytes).Doc.Content![0]);

        Assert.Contains(depois, item => TextOfFloat(item) == "Título trocado");
        Assert.Contains(depois, item => TextOfFloat(item) == "Subtítulo do manual");
        Assert.DoesNotContain(depois, item => TextOfFloat(item) == "Título do manual");

        // O ramo de reserva também, senão o arquivo diz duas coisas.
        Assert.DoesNotContain("Título do manual", MainDocumentXml(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CaixaNaoEditadaNaoEReescrita()
    {
        // A caixa em que ninguém tocou segue com o XML que tinha.
        var original = Fixtures.WithTextBoxes();
        var model = Clone(Open(original));

        var capa = model.Doc.Content![0];
        capa.Content = [new Node { Type = "text", Text = "Capa" }];

        var xml = MainDocumentXml(Save(original, model).Bytes);

        Assert.Contains("<wps:bodyPr", xml, StringComparison.Ordinal);
        Assert.Contains("Subtítulo do manual", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void CaixaDeTextoNaoTravaMaisODocumento()
    {
        // Editável, porque editar a âncora copia a forma. O aviso fica: sem
        // declaração, a forma herda moldura e preenchimento do tema.
        var result = DocxReader.Read(Fixtures.WithTextBoxes());

        Assert.Contains(Inventory.Shapes, result.Inventory.Invisible);
        Assert.DoesNotContain(Inventory.Shapes, result.Inventory.Structural);
    }

    [Fact]
    public void CaixaSemMolduraNaoAvisaMolduraNenhuma()
    {
        // Caixa sem decoração não pede aviso.
        var result = DocxReader.Read(Fixtures.WithDecoratedTextBoxes());

        Assert.DoesNotContain(Inventory.Shapes, result.Inventory.Invisible);

        var floats = result.Model.Doc.Content!.SelectMany(FloatsOf).ToList();

        Assert.Null(TextOf(floats[0], "fill"));
        Assert.Null(TextOf(floats[0], "line"));

        // 12700 EMU são 1 pt.
        Assert.Equal("#ffffff", TextOf(floats[1], "fill"));
        Assert.Equal("#1f5fa9", TextOf(floats[1], "line"));
        Assert.Equal(1, floats[1].GetProperty("lineWidthPt").GetDouble());
        Assert.False(floats[1].GetProperty("dash").GetBoolean());
    }

    [Fact]
    public void OGradienteContinuaSendoAvisado()
    {
        // O aviso passa a falar só do que sobra: gradiente, textura, sombra, canto arredondado.
        var result = DocxReader.Read(Fixtures.WithGradientTextBox());

        Assert.Contains(Inventory.Shapes, result.Inventory.Invisible);
    }

    [Fact]
    public void OCabecalhoTrazAFonteQueODocumentoPede()
    {
        // Sem ela o cabeçalho herdaria a fonte do editor.
        var floats = Open(Fixtures.WithHeaderGroup()).Page.Header!.Floats;
        var titulo = floats.Single(item => item.Kind == "text");

        var marca = titulo.Content!.Single().Content!.Single().Marks!
            .Single(m => m.Type == "textStyle");

        // Sem pilha: o fixture não declara a família, e a `@font-face` do editor serve a Calibri.
        Assert.Equal("Calibri", marca.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void OFileteDoCabecalhoEDesenhado()
    {
        // O filete é uma forma de altura zero com contorno, no grupo do logotipo.
        var floats = Open(Fixtures.WithHeaderGroup()).Page.Header!.Floats;

        var filete = Assert.Single(floats.Where(item => item.Kind == "rule"));
        Assert.True(filete.WidthMm > 100);
        Assert.Equal(0, filete.HeightMm);
    }

    [Fact]
    public void AListaTrazAMarcaEOsRecuosQueODocumentoPede()
    {
        // O documento diz o caractere e as duas distâncias, onde o texto começa e quanto
        // o marcador fica antes dele; o caractere da área de uso privado vem traduzido.
        var model = Open(Fixtures.WithBulletList());
        var lista = model.Doc.Content!.Single(node => node.Type == "bulletList");

        Assert.Equal("▪", lista.Attrs!["marker"]!.GetValue<string>());
        Assert.Equal(12.7, lista.Attrs["indentMm"]!.GetValue<double>(), 1);
        Assert.Equal(6.35, lista.Attrs["hangingMm"]!.GetValue<double>(), 2);
    }

    [Fact]
    public void CadaParagrafoDoRodapeAbreUmaLinha()
    {
        // Cada parágrafo da faixa é uma linha.
        var page = Open(Fixtures.WithFooterOfThreeLines()).Page;

        var pieces = page.Footer!.Center;
        Assert.Equal(3, pieces.Count);
        Assert.False(pieces[0].Line);
        Assert.True(pieces[1].Line);
        Assert.True(pieces[2].Line);
    }

    [Fact]
    public void CadaPecaDaFaixaSabeDeOndeVeio()
    {
        // O endereço deixa a gravação escrever no `w:t` da peça sem regerar a faixa.
        var pieces = Open(Fixtures.WithFooterOfThreeLines()).Page.Footer!.Center;

        Assert.Equal(3, pieces.Count);
        Assert.Equal(3, pieces.Select(piece => piece.Pid).Distinct().Count());

        foreach (var (piece, at) in pieces.Select((piece, at) => (piece, at)))
        {
            Assert.EndsWith($":{at}:0", piece.Pid);
        }
    }

    [Fact]
    public void PecaSemTextoProprioNaoRecebeEndereco()
    {
        // Sem `w:t`, sem endereço: digitar ali apagaria o campo.
        var pieces = Open(Fixtures.WithFooterOfPageNumber()).Page.Footer!.Left;

        var text = Assert.Single(pieces, piece => piece.Kind == PieceDto.KindText);
        Assert.Equal("Página  ", text.Text);
        Assert.Null(text.Pid);

        var number = Assert.Single(pieces, piece => piece.Kind == PieceDto.KindPageNumber);
        Assert.Null(number.Pid);
    }

    [Fact]
    public void OTextoDigitadoNoRodapeVoltaParaOArquivo()
    {
        // Só o `w:t` da peça editada muda.
        var original = Fixtures.WithFooterOfThreeLines();
        var model = Clone(Open(original));

        var footer = model.Page.Footer!;
        footer.Center[1] = footer.Center[1] with { Text = "Documento V02 - Desenvolvido por: Sicrano" };

        var reopened = Open(Save(original, model).Bytes).Page.Footer!;

        Assert.Equal("www.exemplo.com.br", reopened.Center[0].Text);
        Assert.Equal("Documento V02 - Desenvolvido por: Sicrano", reopened.Center[1].Text);
        Assert.Equal("Mês/ANO", reopened.Center[2].Text);
    }

    [Fact]
    public void RodapeNaoEditadoNaoEReescrito()
    {
        // A parte que ninguém tocou volta byte a byte, e não reserializada pelo SDK.
        var original = Fixtures.WithFooterOfThreeLines();
        var model = Clone(Open(original));

        Assert.Equal(
            PartsOf(original)["word/footer1.xml"],
            PartsOf(Save(original, model).Bytes)["word/footer1.xml"]);
    }

    [Fact]
    public void OCampoDoRodapeSobreviveAEdicaoDaLinha()
    {
        // O texto fundido com a tabulação perdeu o rastro: a gravação não escreve nada,
        // antes que o campo vire número fixo.
        var original = Fixtures.WithFooterOfPageNumber();
        var model = Clone(Open(original));

        var footer = model.Page.Footer!;
        footer.Left[0] = footer.Left[0] with { Text = "Folha " };

        var saved = Save(original, model).Bytes;

        Assert.Equal(PartsOf(original)["word/footer1.xml"], PartsOf(saved)["word/footer1.xml"]);
        Assert.Contains(
            Open(saved).Page.Footer!.Left,
            piece => piece.Kind == PieceDto.KindPageNumber);
    }

    [Fact]
    public void OTextoDigitadoNaCaixaDoCabecalhoVoltaParaOArquivo()
    {
        // O título do cabeçalho corporativo mora numa caixa dentro de um grupo de formas.
        var original = Fixtures.WithHeaderGroup();
        var model = Clone(Open(original));

        var caixa = Assert.Single(model.Page.Header!.Floats!, item => item.Kind == "text");
        Assert.NotNull(caixa.BoxId);

        caixa.Content![0].Content![0].Text = "EVIDÊNCIAS DE HOMOLOGAÇÃO";

        var reopened = Open(Save(original, model).Bytes).Page.Header!;
        var voltou = Assert.Single(reopened.Floats!, item => item.Kind == "text");

        Assert.Equal("EVIDÊNCIAS DE HOMOLOGAÇÃO", voltou.Content![0].Content![0].Text);

        // O grupo segue inteiro: logotipo e filete o escritor não gera.
        Assert.Contains(reopened.Floats!, item => item.Kind == "image");
        Assert.Contains(reopened.Floats!, item => item.Kind == "rule");
    }

    [Fact]
    public void CaixaDoCabecalhoNaoEditadaNaoEReescrita()
    {
        var original = Fixtures.WithHeaderGroup();

        Assert.Equal(
            PartsOf(original)["word/header1.xml"],
            PartsOf(Save(original, Clone(Open(original))).Bytes)["word/header1.xml"]);
    }

    [Fact]
    public void OEspacoDeclaradoNaoApagaAEntrelinhaDoEstilo()
    {
        // Os atributos de `w:spacing` são independentes: declarar só `w:after` não
        // apaga a entrelinha do estilo.
        var blocks = OpenFlat(Fixtures.WithStyleSpacingAndDirectMargins()).Doc.Content!;

        // 276/240 em Arial de 10 pt: 1,15 × a altura natural da fonte.
        Assert.Equal("1.3224", blocks[0].Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal(0d, blocks[0].Attrs!["spaceAfter"]!.GetValue<double>());
    }

    [Fact]
    public void ORecuoDoParagrafoEAMedidaQueODocumentoPede()
    {
        // O nível do editor vale 2,5em, 25 pt a 10 pt, e não os 36 pt de 720 twips.
        var blocks = OpenFlat(Fixtures.WithStyleSpacingAndDirectMargins()).Doc.Content!;

        // 720 twips = 12,7 mm; o pendente vem negativo, porque é `text-indent`.
        Assert.Equal(12.7, blocks[0].Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(1.06, blocks[0].Attrs!["indentRightMm"]!.GetValue<double>());
        Assert.Equal(-6.35, blocks[0].Attrs!["firstLineMm"]!.GetValue<double>());
        Assert.Equal(0, blocks[0].Attrs!["indent"]!.GetValue<int>());

        // Trocar o recuo da esquerda não apaga o pendente que veio do estilo.
        Assert.Equal(25.4, blocks[1].Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(-6.35, blocks[1].Attrs!["firstLineMm"]!.GetValue<double>());
    }

    [Fact]
    public void ORecuoVoltaParaOArquivoNaMedidaEmQueVeio()
    {
        var original = Fixtures.WithStyleSpacingAndDirectMargins();
        var model = Clone(Open(original));

        // O bloco editado é reescrito do zero.
        model.Doc.Content![0].Content![0].Text = "Outro texto no mesmo recuo.";

        // Relido achatado, o recuo herdado do estilo continua valendo.
        var reopened = OpenFlat(Save(original, model).Bytes).Doc.Content![0];
        Assert.Equal(12.7, reopened.Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(-6.35, reopened.Attrs!["firstLineMm"]!.GetValue<double>());
    }

    [Fact]
    public void OTrechoLevaSoOQueDifereDoEstilo()
    {
        // O trecho sem formatação própria não leva a do estilo, senão modificar o
        // estilo não mudaria a tela.
        var banner = FirstOfType(Open(Fixtures.WithStyles()), "paragraph");
        Assert.Empty(MarksOf(banner));

        // A leitura achatada, do rascunho antigo, leva tudo.
        var flat = FirstOfType(OpenFlat(Fixtures.WithStyles()), "paragraph");
        Assert.Contains(MarksOf(flat), m => m.Type == "bold");
    }

    [Fact]
    public void OBlocoLevaSoOQueOParagrafoDeclara()
    {
        // O bloco leva só o direto, no valor efetivo; o herdado vem do CSS dos estilos.
        var blocks = Open(Fixtures.WithStyleSpacingAndDirectMargins()).Doc.Content!;

        var first = blocks[0].Attrs!;
        Assert.Equal(0d, first["spaceAfter"]!.GetValue<double>());
        Assert.Equal(0d, first["spaceBefore"]!.GetValue<double>());
        Assert.False(first.ContainsKey("lineHeight"));
        Assert.False(first.ContainsKey("indentMm"));
        Assert.False(first.ContainsKey("firstLineMm"));
        Assert.False(first.ContainsKey("fontFamily"));
        Assert.Equal(0, first["indent"]!.GetValue<int>());

        var moved = blocks[1].Attrs!;
        Assert.Equal(25.4, moved["indentMm"]!.GetValue<double>());
        Assert.False(moved.ContainsKey("firstLineMm"));
    }

    [Fact]
    public void OFundoEOManterComOProximoDoEstiloFicamNoEstilo()
    {
        var model = Open(Fixtures.WithStyles());
        var banner = FirstOfType(model, "paragraph");

        Assert.False(banner.Attrs!.ContainsKey("background"));
        Assert.False(banner.Attrs.ContainsKey("textAlign"));
        Assert.Equal("Faixa", banner.Attrs["styleId"]!.GetValue<string>());
        Assert.Equal("#943634", model.Styles!.Styles["Faixa"].Paragraph!.Background);
    }

    [Fact]
    public void OEditadoGravaAEntrelinhaNaFonteDoEstiloESomaONivelAoRecuoDoEstilo()
    {
        // Fonte e recuo do estilo o gravador busca no estilo.
        var original = Fixtures.WithStyleSpacingAndDirectMargins();
        var model = Clone(Open(original));
        var block = model.Doc.Content![0];
        block.Content![0].Text = "Editado.";
        block.With("lineHeight", "1.7249").With("indent", 1);

        var xml = Roundtrip.XmlOf(Save(original, model).Bytes, "word/document.xml");
        // 1,7249 ÷ 1,1499 (Arial) × 240 = 360.
        Assert.Contains("w:line=\"360\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:left=\"1440\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ORascunhoAchatadoEComparadoComALeituraAchatada()
    {
        // O rascunho de antes do `.sdoc` 4 traz blocos achatados.
        var original = Fixtures.WithStyles();
        var flat = Clone(OpenFlat(original));

        var wrong = Save(original, flat).Result;
        Assert.True(wrong.RewrittenBlocks > 0);

        var right = Save(original, flat with { Flatten = true }).Result;
        Assert.Equal(0, right.RewrittenBlocks);
    }

    [Fact]
    public void AMarcaDeSecaoNaoEUmaLinhaDeTexto()
    {
        // A marca de seção é um parágrafo vazio com `w:sectPr`, e o LibreOffice não
        // lhe dá altura.
        var original = Fixtures.WithSectionMarkInTheMiddle();
        var model = Clone(Open(original));

        var mark = model.Doc.Content![1];
        Assert.True(mark.Attrs!["sectionMark"]!.GetValue<bool>());

        // E continua no arquivo, porque carrega a seção.
        var saved = Save(original, model);
        Assert.Equal(3, saved.Result.PreservedBlocks);
        Assert.True(Open(saved.Bytes).Doc.Content![1].Attrs!["sectionMark"]!.GetValue<bool>());
    }

    [Fact]
    public void OParagrafoQueAncoraEQuebraContinuaParagrafo()
    {
        // A marca da capa ancorada no parágrafo da quebra fica na folha de cima: o nó
        // de quebra mora entre as folhas.
        var model = Open(Fixtures.WithBreakOnAnchorParagraph());

        var capa = model.Doc.Content![0];
        Assert.Equal("paragraph", capa.Type);
        Assert.True(capa.Attrs!["breakAfter"]!.GetValue<bool>());
        Assert.Single(FloatsOf(capa));
    }

    [Fact]
    public void QuebraSozinhaContinuaSendoNoDeQuebra()
    {
        // Sem nada ancorado, o parágrafo só com a quebra é o nó de quebra.
        var model = Open(Fixtures.WithLonePageBreak());

        Assert.Contains(model.Doc.Content!, node => node.Type == "pageBreak");
    }

    [Fact]
    public void OGiroViajaEmGrausSemMexerNasMedidas()
    {
        // O Word posiciona sem girar e gira em torno do centro, como `transform: rotate()`.
        var model = Open(Fixtures.WithRotatedImage());

        var image = Assert.Single(FloatsOf(model.Doc.Content![0]));
        Assert.Equal(270, image.GetProperty("rotation").GetDouble());
        Assert.Equal(285.76, image.GetProperty("widthMm").GetDouble(), 2);
        Assert.Equal(80.14, image.GetProperty("heightMm").GetDouble(), 2);
    }

    [Fact]
    public void AAncoraViajaComOrigemEDeslocamento()
    {
        // A origem mais comum é o parágrafo, que só tem posição depois de paginar.
        var image = Assert.Single(FloatsOf(Open(Fixtures.WithRotatedImage()).Doc.Content![0]));

        Assert.Equal("column", image.GetProperty("hFrom").GetString());
        Assert.Equal("paragraph", image.GetProperty("vFrom").GetString());
    }

    [Theory]
    [InlineData("não é um zip")]
    [InlineData("PK zip truncado")]
    public void RejectsGarbageWithAReadableMessage(string garbage)
    {
        // Documento é dado não confiável: vira frase, e nunca derruba o processo.
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

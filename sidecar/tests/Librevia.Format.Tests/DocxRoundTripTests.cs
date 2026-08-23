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
    public void ImagemAncoradaSaiDoFluxo()
    {
        // No Word um objeto ancorado não empurra o texto: mora numa posição da
        // folha. Lido como bloco, a marca vertical de 28,6 cm da capa entrava
        // como faixa deitada de página inteira e empurrava tudo para baixo — a
        // contagem de páginas ia junto.
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
    public void ImagemNoFluxoContinuaSendoBloco()
    {
        // O contrapeso: `wp:inline` é imagem no meio da linha, e deve mesmo
        // ocupar lugar. Tirar as duas do fluxo esvaziaria o documento.
        var model = Open(Fixtures.WithInlineImage());

        var image = Assert.Single(Walk(model.Doc).Where(n => n.Type == "image"));
        Assert.Equal(554, image.Attrs!["width"]!.GetValue<int>());

        // A altura junto com a largura. Sem ela o navegador reserva zero até os
        // bytes decodificarem — e a paginação mede a folha uma vez, antes disso.
        Assert.Equal(277, image.Attrs["height"]!.GetValue<int>());
    }

    [Fact]
    public void ImagemReescritaMantemOTamanhoQueTinha()
    {
        // O parágrafo da imagem foi editado, então ele é regravado do modelo em
        // vez de copiado. Enquanto a altura não chegava, o escritor chutava três
        // quartos da largura: a imagem voltava para o arquivo com outra forma,
        // sem ninguém ter tocado nela.
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
        // No arquivo a lista não existe como bloco: são parágrafos numerados,
        // cada um com o seu espaçamento. Na árvore do editor ela é um elemento
        // de verdade, e elemento sem espaçamento declarado recebe o do editor —
        // o mesmo que o parágrafo importado já não recebe. Num documento com
        // seis listas isso somava quinze milímetros de ar que o Word não tem.
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
    public void ParagraphWithoutAStyleStillGetsTheDefaultOne()
    {
        // O estilo marcado `w:default="1"` vale para quem não declara
        // `w:pStyle`. Sem consultá-lo, o texto do corpo de um documento inteiro
        // chegava só com os `docDefaults` — quer dizer, quase sem formatação.
        var model = Open(Fixtures.WithLineMetrics());
        var first = BlockContaining(model, "estilo padrão");

        Assert.Equal("12pt", first.Attrs!["fontSize"]!.GetValue<string>());
    }

    [Fact]
    public void TheParagraphMarkGivesTheBlockItsFont()
    {
        // A marca de parágrafo continua fora dos runs — pô-la ali punha negrito
        // em parágrafo que não tem. Mas é com ela que o Word mede a linha, e o
        // bloco precisa dela: um parágrafo vazio de Verdana 10 pt ocupava os
        // 12 pt do padrão do editor, meia linha a mais, em cada parágrafo.
        var model = Open(Fixtures.WithLineMetrics());
        var marked = BlockContaining(model, "Verdana");

        Assert.Equal("Verdana", marked.Attrs!["fontFamily"]!.GetValue<string>());
        Assert.Equal("10pt", marked.Attrs["fontSize"]!.GetValue<string>());

        // E continua fora dos runs: o texto segue com a fonte do estilo.
        var run = MarksOf(marked).Single(m => m.Type == "textStyle");
        Assert.Equal("Times New Roman", run.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void SilenceAboutSpacingMeansSingleAndZero()
    {
        // Arquivo calado não está pedindo o padrão de quem o abre: está dizendo
        // zero de espaço e espaçamento simples. Enquanto os atributos ficavam
        // ausentes, o `margin-top` e o `line-height` que o editor traz para o
        // documento em branco reapareciam em cada parágrafo importado — mais de
        // uma página de ar num documento de meia centena de parágrafos.
        var model = Open(Fixtures.WithLineMetrics());
        var first = BlockContaining(model, "estilo padrão");

        Assert.Equal("normal", first.Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal(0, first.Attrs["spaceBefore"]!.GetValue<double>());
        Assert.Equal(0, first.Attrs["spaceAfter"]!.GetValue<double>());
    }

    [Fact]
    public void ReadsLineSpacingThatIsLockedInPoints()
    {
        // `exact` e `atLeast` dizem a altura em twips. As duas voltavam nulas, e
        // um parágrafo com entrelinha travada era desenhado com a do editor.
        var model = Open(Fixtures.WithLineMetrics());

        Assert.Equal("9pt", BlockContaining(model, "travada").Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal("1.5", BlockContaining(model, "uma vez e meia").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void AFonteAusenteCaiNaSubstitutaDoTipoCerto()
    {
        // O documento nomeia fontes que a máquina de quem abre pode não ter. Sem
        // dizer de que tipo elas são, o navegador cai na próxima da pilha, e a
        // pilha do editor termina em serifa: a capa do modelo de manual, que
        // pede Segoe UI, saía com o título em Times enquanto o LibreOffice o
        // desenha sem serifa. `word/fontTable.xml` é quem diz o tipo.
        var model = Open(Fixtures.WithMissingFont());

        var titulo = MarksOf(BlockContaining(model, "Título da capa")).Single(m => m.Type == "textStyle");
        Assert.Equal("Segoe UI, sans-serif", titulo.Attrs!["fontFamily"]!.GetValue<string>());

        // Fonte que a tabela não declara não ganha pilha inventada.
        var outro = MarksOf(BlockContaining(model, "Sem tipo")).Single(m => m.Type == "textStyle");
        Assert.Equal("Fonte Fantasma", outro.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void SoONomeDaFonteVoltaParaOArquivo()
    {
        // `w:rFonts` guarda o nome de uma fonte, não uma pilha de CSS. Escrever
        // a pilha inteira gravaria uma fonte chamada "Segoe UI, sans-serif".
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
        // Diz que o bloco não fica sozinho no pé da página. A marca de fim de
        // página e a folha de impressão usam o mesmo sinal — se divergissem, a
        // marca cairia num lugar e o PDF quebraria noutro.
        var model = Open(Fixtures.WithKeepNext());
        var kept = BlockContaining(model, "Rótulo");

        Assert.True(kept.Attrs!["keepNext"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Solto").Attrs?.ContainsKey("keepNext") ?? false);
    }

    [Fact]
    public void TreatsLeadingTabsAsCentering()
    {
        // No corpus, o primeiro título de cada documento vem alinhado à
        // esquerda com tabulações até uma parada centralizada. No Word ele
        // aparece no meio; no HTML a tabulação colapsa e o título encostava à
        // esquerda, destoando dos títulos vizinhos.
        var model = Open(Fixtures.WithTabCentering());
        var centered = BlockContaining(model, "Centralizado por tabulação");

        Assert.Equal("center", centered.Attrs!["textAlign"]!.GetValue<string>());

        // As tabulações saem **deste** parágrafo: viraram posicionamento, não
        // conteúdo. A do outro parágrafo, no meio da linha, continua lá.
        var text = string.Concat(Walk(centered).Where(n => n.Type == "text").Select(n => n.Text));
        Assert.Equal("Centralizado por tabulação", text);
    }

    [Fact]
    public void LeavesTabsInTheMiddleOfALineAlone()
    {
        // A regra só vale para tabulação no começo da linha. "Esquerda [tab]
        // meio" é outra coisa, e centralizar o parágrafo inteiro o quebraria.
        var model = Open(Fixtures.WithTabCentering());
        var inline = BlockContaining(model, "Esquerda");

        Assert.NotEqual("center", inline.Attrs?["textAlign"]?.GetValue<string>());
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

    [Fact]
    public void ComentarioEControleDeAlteracoesSaoEstruturais()
    {
        // É esta lista que decide se o documento abre em somente leitura. Ela é
        // subconjunto da invisibilidade: o recurso continua no arquivo, e só
        // some se o usuário editar justamente o bloco que o ancora.
        var comentado = DocxReader.Read(Fixtures.WithComment()).Inventory;
        var revisado = DocxReader.Read(Fixtures.WithTrackedChanges()).Inventory;

        Assert.Contains(Inventory.Comments, comentado.Structural);
        Assert.Contains(Inventory.TrackedChanges, revisado.Structural);
        // Subconjunto, não lista paralela: tudo que é estrutural também é
        // invisível, senão o aviso de tela deixaria de mencioná-lo.
        Assert.All(comentado.Structural, item => Assert.Contains(item, comentado.Invisible));
    }

    [Fact]
    public void DocumentoComumNaoAbreEmSomenteLeitura()
    {
        // Somente leitura por padrão em documento normal travaria o uso do dia
        // a dia — e o usuário aprenderia a clicar "editar mesmo assim" sem ler.
        var result = DocxReader.Read(Fixtures.Simple());

        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void ImagemAncoradaNaoEEstrutural()
    {
        // Perda de **aparência**, não de conteúdo: a imagem aparece na tela e
        // volta para o arquivo. Travar a edição por causa dela seria o mesmo
        // que travar por nada, já que quase todo documento do corpus tem uma.
        var result = DocxReader.Read(Fixtures.WithAnchoredImage());

        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void DesenhoAncoradoNoCabecalhoNaoEspremeEmTresColunas()
    {
        // A faixa é uma grade de três colunas, e é o bastante para texto. Um
        // desenho ancorado traz posição de verdade e pode vir girado: a marca
        // lateral do corpus é uma faixa de 28,6 mm **em pé**, que não entra numa
        // banda de 10 mm de altura — espremida ali, era desenhada deitada e no
        // terço errado.
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
        // É como uma marca sai para fora da coluna pela esquerda. Lida sem
        // sinal, ela atravessaria para o outro lado da página.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(-1123950)).Model.Page;

        var marca = Assert.Single(page.Header!.Floats!);
        Assert.Equal(-31.22, marca.HorizontalOffsetMm!.Value, 2);
    }

    [Fact]
    public void CabecalhoEmGradeSaiComoGrade()
    {
        // Três colunas bastam para texto e não bastam para isto: o cabeçalho
        // corporativo do corpus é uma tabela com o logotipo mesclado por três
        // linhas, o título ao lado e a numeração à direita. Achatado nos terços,
        // virava uma fileira de palavras por cima da primeira linha do texto.
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
        // No OOXML a mesclagem vertical não é altura: a célula de cima diz
        // `restart` e as de baixo aparecem como células vazias. Desenhadas como
        // células de verdade, elas abririam uma faixa em branco por linha
        // debaixo do logotipo.
        var rows = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page.Header!.Rows!;

        Assert.Equal(3, rows[0].Cells[0].RowSpan);
        Assert.Equal("image", rows[0].Cells[0].Pieces[0].Kind);

        // E as linhas de baixo não têm mais a célula do logotipo.
        Assert.Equal(2, rows[1].Cells.Count);
        Assert.Equal(3, rows[2].Cells.Count);
    }

    [Fact]
    public void AGradeTrazLarguraJuntoEBordaResolvida()
    {
        // A largura é fração da grade, e não twips: quem desenha não sabe quanto
        // vale um twip na folha. A borda vem resolvida porque no OOXML cada lado
        // sai de três lugares — a célula, a moldura, a linha interna — e refazer
        // essa conta na tela e outra vez no papel é como os dois divergem.
        var rows = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page.Header!.Rows!;

        Assert.Equal(0.2, rows[0].Cells[0].Width, 3);
        Assert.Equal(2, rows[0].Cells[2].Span);

        Assert.Contains("b", rows[0].Cells[1].Borders, StringComparison.Ordinal);
        // `w:nil` na célula apaga o risco que a tabela pediu.
        Assert.DoesNotContain("b", rows[2].Cells[1].Borders, StringComparison.Ordinal);
    }

    [Fact]
    public void ADistanciaDaFaixaAteABordaEhLida()
    {
        // É a origem vertical das âncoras de dentro da faixa: elas se dizem
        // relativas ao parágrafo, e o parágrafo do cabeçalho começa aqui. Sem
        // esta medida não há de onde contar.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(0)).Model.Page;

        Assert.Equal(12.5, page.HeaderDistanceMm, 1);
    }

    // --- quebra de página ---------------------------------------------------

    [Fact]
    public void QuebraDentroDoParagrafoViraPropriedadeDoBloco()
    {
        // Emiti-la como nó ali dentro poria um bloco em posição de linha, que o
        // editor não aceita — e, serializado, um `<div>` dentro de `<p>` faz o
        // analisador de HTML desalojar o `div`: 15 blocos viram 17 elementos, os
        // índices deixam de casar e o papel corta noutro lugar que a tela.
        var model = Open(Fixtures.WithBreakInsideParagraph());

        var blocos = model.Doc.Content!;
        Assert.DoesNotContain(Walk(model.Doc), n => n.Type == "pageBreak");
        Assert.True(blocos[0].Attrs!["breakAfter"]!.GetValue<bool>());
    }

    [Fact]
    public void AQuebraDoParagrafoVoltaAoArquivoAoEditar()
    {
        // O parágrafo intocado volta pelo XML original e a quebra vem junto. O
        // caso que precisa de gravador é o editado: sem escrever a quebra de
        // volta, ela sumiria em silêncio.
        var bytes = Fixtures.WithBreakInsideParagraph();
        var model = Clone(Open(bytes));
        Assert.True(EditFirstTextContaining(model, "Fim da", "Outro texto."));

        var reaberto = Open(Save(bytes, model).Bytes);

        Assert.Contains("Outro texto.", TextOf(reaberto), StringComparison.Ordinal);
        Assert.True(reaberto.Doc.Content![0].Attrs!["breakAfter"]!.GetValue<bool>());
    }

    // --- cabeçalho por página -----------------------------------------------

    [Fact]
    public void OCabecalhoPadraoNaoDependeDaOrdemDeGravacao()
    {
        // O `w:type` era ignorado: percorria-se as referências na ordem do XML e
        // ficava a primeira não vazia. Neste fixture o `first` vem antes, então
        // o cabeçalho da capa apareceria no documento inteiro — e qual apareceria
        // dependia de como o Word gravou, não do que o documento diz.
        var bytes = Fixtures.WithFirstPageHeader(titlePage: true);

        // A armadilha, afirmada aqui para não depender de quem lê o fixture: a
        // referência `first` vem **antes** da `default` no XML, e não está
        // vazia. Pela ordem, "Capa" ganharia.
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
        // O Word guarda a parte `first` mesmo com o interruptor desligado —
        // quatro dos seis documentos do corpus são assim. Usá-la sem conferir
        // poria a capa em todas as páginas.
        var page = DocxReader.Read(Fixtures.WithFirstPageHeader(titlePage: false)).Model.Page;

        Assert.Null(page.FirstHeader);
        Assert.Equal("Miolo", TextOfBand(page.Header));
    }

    // --- impressão digital --------------------------------------------------

    // O leitor e o editor descrevem o mesmo bloco de formas diferentes. Cada
    // teste abaixo fixa uma dessas diferenças como "não é edição" — sem isso a
    // gravação cirúrgica regenera o documento inteiro achando que o usuário
    // mexeu em tudo. Medido: `modelo-de-manual.docx` aberto e salvo sem
    // edição nenhuma preservava 2 dos 15 blocos.

    [Fact]
    public void AtributoNuloEAusenteSaoAMesmaCoisa()
    {
        // O ProseMirror materializa todo atributo do schema, inclusive os que o
        // documento não menciona, e devolve `null` neles.
        var doLeitor = Node.Of("paragraph").With("fontSize", "12pt");
        var doEditor = Node.Of("paragraph").With("fontSize", "12pt").With("styleId", null);

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void AOrdemDasChavesNaoContaComoEdicao()
    {
        // `JsonObject` guarda a ordem de inserção, e os dois lados montam os
        // atributos em ordens diferentes. Sozinha, esta diferença reprovava
        // todos os blocos — foi a última a aparecer e a que escondia as outras.
        var doLeitor = Node.Of("paragraph").With("lineHeight", 1.16).With("fontSize", "12pt");
        var doEditor = Node.Of("paragraph").With("fontSize", "12pt").With("lineHeight", 1.16);

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void AOrdemDasMarcasNaoContaComoEdicao()
    {
        // Negrito antes ou depois da cor é a mesma formatação; o ProseMirror
        // ordena as marcas pela posição delas no schema.
        var doLeitor = new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("bold"), Mark.Of("textStyle", "color", "#404040")] };
        var doEditor = new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("textStyle", "color", "#404040"), Mark.Of("bold")] };

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void TextoVizinhoComAsMesmasMarcasEUmTextoSo()
    {
        // O leitor emite um nó por `w:r`: "Acme® Software" chega partido
        // quando o `®` tem `rPr` próprio, ainda que igual ao do vizinho.
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
        // O contrapeso: fundir sem olhar as marcas apagaria a diferença entre
        // "Acme" em negrito e "® Software" sem negrito, e a comparação
        // deixaria de perceber uma edição de formatação de verdade.
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
        // Toda normalização acima afrouxa a comparação. Se afrouxar demais, a
        // gravação preserva o XML antigo de um bloco que o usuário mudou — que
        // é perda de dados, e pior que regenerar.
        var antes = Node.Of("paragraph", new Node { Type = "text", Text = "Sumário" });
        var depois = Node.Of("paragraph", new Node { Type = "text", Text = "Sumario" });

        Assert.NotEqual(antes.Fingerprint(), depois.Fingerprint());
    }

    [Fact]
    public void OidNaoEntraNaComparacao()
    {
        // Identidade não é conteúdo: dois blocos iguais em posições diferentes
        // do documento têm `oid` diferente e o mesmo texto.
        var primeiro = Node.Of("paragraph", new Node { Type = "text", Text = "Texto" }).With("oid", "b1");
        var segundo = Node.Of("paragraph", new Node { Type = "text", Text = "Texto" }).With("oid", "b9");

        Assert.Equal(primeiro.Fingerprint(), segundo.Fingerprint());
    }

    // --- formas e caixas de texto -------------------------------------------

    [Fact]
    public void CaixaDeTextoViraObjetoComOTextoDentro()
    {
        // Uma caixa é um fluxo de texto próprio, posicionado na folha. Antes ela
        // ia inteira para o inventário e sumia da tela; depois passou a ter o
        // texto despejado na linha da âncora, o que mostrava o conteúdo mas
        // emendava título e subtítulo numa frase só.
        var floats = FloatsOf(Open(Fixtures.WithTextBoxes()).Doc.Content![0]);

        Assert.Equal(2, floats.Count);
        Assert.All(floats, item => Assert.Equal("text", item.GetProperty("kind").GetString()));
        Assert.Contains(floats, item => TextOfFloat(item) == "Título do manual");
        Assert.Contains(floats, item => TextOfFloat(item) == "Subtítulo do manual");
    }

    [Fact]
    public void CaixaDeTextoNaoEntraDuasVezes()
    {
        // O Word grava a mesma caixa em DrawingML e no VML de reserva. Ler os
        // dois ramos criaria dois objetos no mesmo lugar, um sobre o outro.
        var floats = FloatsOf(Open(Fixtures.WithTextBoxes()).Doc.Content![0]);

        Assert.Single(floats.Where(item => TextOfFloat(item) == "Título do manual"));
    }

    [Fact]
    public void CaixaDeTextoContinuaSendoEstrutural()
    {
        // Mostrar o texto não devolve a forma: a moldura e a posição continuam
        // invisíveis, e a gravação cirúrgica ainda perde a caixa inteira se o
        // parágrafo âncora for editado. É o que mantém o documento travado em
        // somente leitura, e ler o texto não pode afrouxar isso.
        var result = DocxReader.Read(Fixtures.WithTextBoxes());

        Assert.Contains(Inventory.Shapes, result.Inventory.Structural);
    }

    [Fact]
    public void OGiroViajaEmGrausSemMexerNasMedidas()
    {
        // O Word posiciona a caixa sem girar e depois a gira em torno do centro
        // — o mesmo que `transform: rotate()`. Trocar largura por altura aqui
        // deslocaria o objeto por metade da diferença entre as duas.
        var model = Open(Fixtures.WithRotatedImage());

        var image = Assert.Single(FloatsOf(model.Doc.Content![0]));
        Assert.Equal(270, image.GetProperty("rotation").GetDouble());
        Assert.Equal(285.76, image.GetProperty("widthMm").GetDouble(), 2);
        Assert.Equal(80.14, image.GetProperty("heightMm").GetDouble(), 2);
    }

    [Fact]
    public void AAncoraViajaComOrigemEDeslocamento()
    {
        // Não são resolvidos aqui: a origem vertical mais comum é o parágrafo, e
        // parágrafo só tem posição depois de paginar. Quem desenha faz a conta.
        var image = Assert.Single(FloatsOf(Open(Fixtures.WithRotatedImage()).Doc.Content![0]));

        Assert.Equal("column", image.GetProperty("hFrom").GetString());
        Assert.Equal("paragraph", image.GetProperty("vFrom").GetString());
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

using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Nodes;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// The surgical save's promise: editing a document must not cost what was not edited.
/// </summary>
public class DocxRoundTripTests
{
    // The path steps live in Roundtrip, shared with DocxWriteBackTests.
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

        // The anchor must stay in the body, or the comment becomes orphaned.
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
        // Opening and saving without changes, the most common case, costs nothing.
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
        // The anchor comes back with the paragraph; see CommentsTests.
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
        // In Word an anchored object lives at a sheet position and does not push the text.
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
        // LibreOffice writes the image in its own paragraph as a `wp:anchor` without offset,
        // centered: it takes height in the flow.
        var model = Open(Fixtures.WithAnchoredImageInTheFlow());

        var image = Assert.Single(Walk(model.Doc).Where(n => n.Type == "image"));
        Assert.Equal(554, image.Attrs!["width"]!.GetValue<int>());

        Assert.All(model.Doc.Content!, block => Assert.Empty(FloatsOf(block)));
    }

    [Fact]
    public void ImagemNoFluxoContinuaSendoBloco()
    {
        // `wp:inline` takes room in the line.
        var model = Open(Fixtures.WithInlineImage());

        var image = Assert.Single(Walk(model.Doc).Where(n => n.Type == "image"));
        Assert.Equal(554, image.Attrs!["width"]!.GetValue<int>());

        // Without the height the browser reserves zero until decoding, and pagination measures
        // before.
        Assert.Equal(277, image.Attrs["height"]!.GetValue<int>());
    }

    [Fact]
    public void ImagemReescritaMantemOTamanhoQueTinha()
    {
        // The edited paragraph is rewritten from the model, and the image keeps its shape.
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
        // In the file the list is numbered paragraphs, each with its own spacing; the editor's list
        // element must not add the default spacing.
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
        // Reading a new feature must not make the model drift from the file, or opening and saving
        // rewrites the paragraph.
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
        // Column width comes from `w:tblGrid`, which TableKit resizes.
        var table = Walk(Open(Fixtures.WithStyledTable()).Doc).First(n => n.Type == "table");
        var header = table.Content![0].Content![0];

        // `w:tblHeader` is the row Word repeats at the top of each page.
        Assert.Equal("tableHeader", header.Type);
        Assert.Equal("tableCell", table.Content[1].Content![0].Type);

        // 15 twips per pixel: 4000 is 267 px.
        Assert.Equal("[267]", header.Attrs!["colwidth"]!.ToJsonString());
        Assert.Equal("[333]", table.Content[0].Content![1].Attrs!["colwidth"]!.ToJsonString());
        Assert.Equal("#d9d9d9", header.Attrs["shading"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsCellBordersInTheCanonicalForm()
    {
        // A single spelling, compared with the editor's: side, style, width in points and color.
        var model = Open(Fixtures.WithPatternedCell());
        var cell = Walk(model.Doc).First(n => n.Type == "tableCell");

        // 24 eighths of a point is 3 pt; `thickThinSmallGap` becomes a single line, and the pattern
        // a flat color.
        Assert.Equal("top:single,3,#000000", cell.Attrs!["borders"]!.GetValue<string>());
        Assert.Equal("#ffff00", cell.Attrs["shading"]!.GetValue<string>());
    }

    [Fact]
    public void ReadsImageAlternativeText()
    {
        // `wp:docPr/@descr` is what the screen reader reads instead of the image.
        var image = Walk(Open(Fixtures.WithDescribedImage()).Doc).First(n => n.Type == "image");

        Assert.Equal("Organograma da diretoria", image.Attrs!["alt"]!.GetValue<string>());
    }

    [Fact]
    public void FlattensIdenticalSectionsWithoutReportingLoss()
    {
        // Seven identical sections are a LibreOffice artifact, and need no warning.
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

    /// <summary>`word/document.xml` from inside the package.</summary>
    private static string MainDocumentXml(byte[] bytes)
    {
        using var buffer = new MemoryStream(bytes, writable: false);
        using var package = new System.IO.Compression.ZipArchive(buffer);
        using var stream = package.GetEntry("word/document.xml")!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>A block's anchored objects.</summary>
    private static List<JsonElement> FloatsOf(Node node)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue("floats", out var value) || value is null)
        {
            return [];
        }

        return [.. JsonDocument.Parse(value.ToJsonString()).RootElement.EnumerateArray()];
    }

    /// <summary>A text property of the object, or null when absent.</summary>
    private static string? TextOf(JsonElement item, string name) =>
        item.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    /// <summary>All the text inside a box.</summary>
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

    /// <summary>All the text of a band, across the three columns.</summary>
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
    /// The first block with the text. Filters by type because <c>Walk</c> includes the root, which
    /// would match any search.
    /// </summary>
    private static Node BlockContaining(DocumentModelDto model, string needle) =>
        Walk(model.Doc)
            .Where(n => n.Type is "paragraph" or "heading")
            .First(n => Walk(n).Any(c => c.Text?.Contains(needle, StringComparison.Ordinal) == true));

    [Fact]
    public void ResolvesFormattingThatLivesOnlyInTheStyle()
    {
        // The corpus `Heading1` is a red bar with white text, all in the style.
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
        // `w:rPr` inside `w:pPr` formats the paragraph mark, not the text.
        var model = Open(Fixtures.WithStyles());
        var paragraph = BlockContaining(model, "negrito");

        Assert.DoesNotContain(MarksOf(paragraph), m => m.Type == "bold");
    }

    [Fact]
    public void PutsTheFontOnTheBlockAndNotOnlyOnTheRuns()
    {
        // Line height comes from the element's font, not from the text inside it.
        var model = OpenFlat(Fixtures.WithStyles());
        var banner = FirstOfType(model, "paragraph");

        Assert.Equal("10pt", banner.Attrs!["fontSize"]!.GetValue<string>());
        Assert.Equal("Arial", banner.Attrs["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void ParagraphWithoutAStyleStillGetsTheDefaultOne()
    {
        // The `w:default="1"` style applies to whoever declares no `w:pStyle`.
        var model = OpenFlat(Fixtures.WithLineMetrics());
        var first = BlockContaining(model, "estilo padrão");

        Assert.Equal("12pt", first.Attrs!["fontSize"]!.GetValue<string>());
    }

    [Fact]
    public void TheParagraphMarkGivesTheBlockItsFont()
    {
        // Word measures the line with the paragraph mark, the empty paragraph's too.
        var model = Open(Fixtures.WithLineMetrics());
        var marked = BlockContaining(model, "Verdana");

        Assert.Equal("Verdana", marked.Attrs!["fontFamily"]!.GetValue<string>());
        Assert.Equal("10pt", marked.Attrs["fontSize"]!.GetValue<string>());

        // And it stays out of the runs.
        Assert.DoesNotContain(MarksOf(marked), m => m.Type == "textStyle");
        var flat = BlockContaining(OpenFlat(Fixtures.WithLineMetrics()), "Verdana");
        var run = MarksOf(flat).Single(m => m.Type == "textStyle");
        Assert.Equal("Times New Roman", run.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void SilenceAboutSpacingMeansSingleAndZero()
    {
        // A silent file means zero space and single spacing, not the editor default.
        var model = OpenFlat(Fixtures.WithLineMetrics());
        var first = BlockContaining(model, "estilo padrão");

        // Single comes out as a number, the font's natural height: Chromium rounds `normal` to a
        // whole pixel.
        Assert.Equal("1.1499", first.Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal(0, first.Attrs["spaceBefore"]!.GetValue<double>());
        Assert.Equal(0, first.Attrs["spaceAfter"]!.GetValue<double>());
    }

    [Fact]
    public void ReadsLineSpacingThatIsLockedInPoints()
    {
        // `exact` and `atLeast` state the height in twips.
        var model = Open(Fixtures.WithLineMetrics());

        Assert.Equal("9pt", BlockContaining(model, "travada").Attrs!["lineHeight"]!.GetValue<string>());

        // The multiple is over the natural height: one and a half Liberation Serif lines is 1.7248
        // em.
        Assert.Equal("1.7248", BlockContaining(model, "uma vez e meia").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void OMultiploDaEntrelinhaEVezAAlturaDaFonte()
    {
        // "1.13 lines" is 1.13 times the font's single line, which in Arial is 1.1499 em.
        var model = Open(Fixtures.WithLineMetrics());

        // 271/240 a quatro casas (1,1292), a grade do arquivo, sobre 1,1499.
        Assert.Equal("1.2985", BlockContaining(model, "Arial e um pouco mais").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void FonteQueNaoVaiNoInstaladorAindaRespeitaOMultiplo()
    {
        // An unknown font with a declared multiple: the 1.15 guess, the height of almost every
        // Latin font, because against the font size it errs by 15 %.
        var model = Open(Fixtures.WithLineMetrics());

        // 271/240 over the 1.15 guess.
        Assert.Equal("1.2985", BlockContaining(model, "fonte que ninguém tem").Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal("normal", BlockContaining(model, "Verdana de dez").Attrs!["lineHeight"]!.GetValue<string>());
    }

    [Fact]
    public void AFonteAusenteCaiNaSubstitutaDoTipoCerto()
    {
        // `word/fontTable.xml` gives the kind of a font that may be missing, for the CSS stack.
        var model = Open(Fixtures.WithMissingFont());

        var titulo = MarksOf(BlockContaining(model, "Título da capa")).Single(m => m.Type == "textStyle");
        Assert.Equal("Segoe UI, sans-serif", titulo.Attrs!["fontFamily"]!.GetValue<string>());

        // Without a declared kind, no invented stack.
        var outro = MarksOf(BlockContaining(model, "Sem tipo")).Single(m => m.Type == "textStyle");
        Assert.Equal("Fonte Fantasma", outro.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void SoONomeDaFonteVoltaParaOArquivo()
    {
        // `w:rFonts` stores the name, not the stack: otherwise the font would be "Segoe UI,
        // sans-serif".
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
        // The page end mark and printing use the same signal, or they would break in different
        // places.
        var model = Open(Fixtures.WithKeepNext());
        var kept = BlockContaining(model, "Rótulo");

        Assert.True(kept.Attrs!["keepNext"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Solto").Attrs?.ContainsKey("keepNext") ?? false);
    }

    [Fact]
    public void PutsTheTopAnchoredImageBeforeTheText()
    {
        // A frame anchored to the top comes before the text, as Word and LibreOffice draw it.
        var model = Open(Fixtures.WithTextAroundTopAnchoredImage());
        var paragraph = model.Doc.Content!.Single(node => node.Type == "paragraph");

        Assert.Equal(["image", "text", "text"], paragraph.Content!.Select(node => node.Type));
    }

    [Fact]
    public void ResolvesCellMarginsLikeWord()
    {
        // Side by side: the table, the default table style, and Word's 0/108.
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
        // On is Word's default: only what turns it off says anything.
        var model = Open(Fixtures.WithKeepLines());

        Assert.False(BlockContaining(model, "Viúva permitida").Attrs!["widowControl"]!.GetValue<bool>());
        Assert.False(BlockContaining(model, "Comum").Attrs?.ContainsKey("widowControl") ?? false);
    }

    [Fact]
    public void TreatsLeadingTabsAsCentering()
    {
        // In the corpus the heading comes left-aligned with tabs up to a centered stop; in HTML the
        // tab collapses.
        var model = Open(Fixtures.WithTabCentering());
        var centered = BlockContaining(model, "Centralizado por tabulação");

        Assert.Equal("center", centered.Attrs!["textAlign"]!.GetValue<string>());

        // This paragraph's tabs become positioning; the one mid-line stays.
        var text = string.Concat(Walk(centered).Where(n => n.Type == "text").Select(n => n.Text));
        Assert.Equal("Centralizado por tabulação", text);
    }

    [Fact]
    public void LeavesTabsInTheMiddleOfALineAlone()
    {
        // Only the tab at the start of the line: "left [tab] middle" is something else.
        var model = Open(Fixtures.WithTabCentering());
        var inline = BlockContaining(model, "Esquerda");

        Assert.NotEqual("center", inline.Attrs?["textAlign"]?.GetValue<string>());
    }

    [Fact]
    public void SurvivesAStyleThatInheritsFromItself()
    {
        // A document is untrusted data: a circular chain does not hang.
        var model = Open(Fixtures.WithCircularStyle());

        Assert.Contains("Texto.", TextOf(model), StringComparison.Ordinal);
    }

    [Fact]
    public void CommentsAreNeitherInvisibleNorLost()
    {
        // The pane shows them: neither an invisibility nor a loss warning.
        var result = DocxReader.Read(Fixtures.WithComment());

        Assert.DoesNotContain("comentários", result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void TrackedChangesNoLongerLockOrWarn()
    {
        // A text revision is an editor mark: neither invisible nor locking.
        var result = DocxReader.Read(Fixtures.WithTrackedChanges());

        Assert.Empty(result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Structural);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void RevisaoDeEstruturaEEstruturalEComentarioNao()
    {
        // The list that decides read-only is a subset of invisibility: comments and text revisions
        // are not in it, structure revisions are.
        var comentado = DocxReader.Read(Fixtures.WithComment()).Inventory;
        var revisado = DocxReader.Read(Fixtures.WithInsertedCell()).Inventory;

        Assert.DoesNotContain(Inventory.Comments, comentado.Structural);
        Assert.Contains(Inventory.StructureRevisions, revisado.Structural);
        // Everything structural is invisible, or the warning would stop mentioning it.
        Assert.All(revisado.Structural, item => Assert.Contains(item, revisado.Invisible));
    }

    [Fact]
    public void DocumentoComumNaoAbreEmSomenteLeitura()
    {
        // Read-only on a normal document would teach people to click "edit anyway" without reading.
        var result = DocxReader.Read(Fixtures.Simple());

        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void ImagemAncoradaNaoEEstrutural()
    {
        // An appearance loss, not a content one: almost every corpus document has one.
        var result = DocxReader.Read(Fixtures.WithAnchoredImage());

        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void DesenhoAncoradoNoCabecalhoNaoEspremeEmTresColunas()
    {
        // An anchored drawing in the band goes as an object, with position and rotation: the corpus
        // side mark stands upright, 28.6 mm in a 10 mm band.
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
        // A mark leaving the column to the left: the sign counts.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(-1123950)).Model.Page;

        var marca = Assert.Single(page.Header!.Floats!);
        Assert.Equal(-31.22, marca.HorizontalOffsetMm!.Value, 2);
    }

    [Fact]
    public void CabecalhoEmGradeSaiComoGrade()
    {
        // The corpus corporate header is a table, with the logo merged across three rows: three
        // columns are not enough.
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
        // OOXML vertical merging is `restart` on top and empty cells below.
        var rows = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page.Header!.Rows!;

        Assert.Equal(3, rows[0].Cells[0].RowSpan);
        Assert.Equal("image", rows[0].Cells[0].Pieces[0].Kind);

        Assert.Equal(2, rows[1].Cells.Count);
        Assert.Equal(3, rows[2].Cells.Count);
    }

    [Fact]
    public void AGradeTrazLarguraJuntoEBordaResolvida()
    {
        // Width as a fraction of the grid, because whoever draws does not know what a twip is
        // worth; the border comes resolved, because each side comes from three places in OOXML.
        var rows = DocxReader.Read(Fixtures.WithHeaderGrid()).Model.Page.Header!.Rows!;

        Assert.Equal(0.2, rows[0].Cells[0].Width, 3);
        Assert.Equal(2, rows[0].Cells[2].Span);

        Assert.Contains("b", rows[0].Cells[1].Borders, StringComparison.Ordinal);
        // `w:nil` on the cell erases the line the table asked for.
        Assert.DoesNotContain("b", rows[2].Cells[1].Borders, StringComparison.Ordinal);
    }

    [Fact]
    public void OGrupoDoCabecalhoSaiPecaPorPeca()
    {
        // The anchor gives the group's place, and `a:chOff`/`a:chExt` the inner ruler.
        var page = DocxReader.Read(Fixtures.WithHeaderGroup()).Model.Page;
        var floats = page.Header!.Floats!;

        // Three pieces: the logo, the title box and the rule below.
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
        // The vertical origin for band anchors, which are relative to the paragraph.
        var page = DocxReader.Read(Fixtures.WithAnchoredHeaderLogo(0)).Model.Page;

        Assert.Equal(12.5, page.HeaderDistanceMm, 1);
    }

    [Fact]
    public void QuebraDentroDoParagrafoViraPropriedadeDoBloco()
    {
        // The editor does not accept a block node inside a line, and HTML moves a `<div>` out of a
        // `<p>`: indexes would stop matching the paper.
        var model = Open(Fixtures.WithBreakInsideParagraph());

        var blocos = model.Doc.Content!;
        Assert.DoesNotContain(Walk(model.Doc), n => n.Type == "pageBreak");
        Assert.True(blocos[0].Attrs!["breakAfter"]!.GetValue<bool>());
    }

    [Fact]
    public void AQuebraDoParagrafoVoltaAoArquivoAoEditar()
    {
        // An edited paragraph must write the break back; an untouched one carries it in the
        // original XML.
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
        // What counts is `w:type`, not the order of references in the XML.
        var bytes = Fixtures.WithFirstPageHeader(titlePage: true);

        // The trap: the non-empty `first` reference comes before `default`.
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
        // Word keeps the `first` part with the switch off (four of the six corpus documents): it
        // comes, switched off, and whoever draws does not use it.
        var page = DocxReader.Read(Fixtures.WithFirstPageHeader(titlePage: false)).Model.Page;

        Assert.False(page.TitlePage);
        Assert.Equal("Capa", TextOfBand(page.FirstHeader));
        Assert.Equal("Miolo", TextOfBand(page.Header));
    }

    // Reader and editor describe the same block differently: each test below pins one of those as
    // "not an edit".
    [Fact]
    public void AtributoNuloEAusenteSaoAMesmaCoisa()
    {
        // ProseMirror materializes every schema attribute, with `null`.
        var doLeitor = Node.Of("paragraph").With("fontSize", "12pt");
        var doEditor = Node.Of("paragraph").With("fontSize", "12pt").With("styleId", null);

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void AOrdemDasChavesNaoContaComoEdicao()
    {
        // `JsonObject` keeps insertion order, which differs between the two sides.
        var doLeitor = Node.Of("paragraph").With("lineHeight", 1.16).With("fontSize", "12pt");
        var doEditor = Node.Of("paragraph").With("fontSize", "12pt").With("lineHeight", 1.16);

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void AOrdemDasMarcasNaoContaComoEdicao()
    {
        // ProseMirror orders marks by their position in the schema.
        var doLeitor = new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("bold"), Mark.Of("textStyle", "color", "#404040")] };
        var doEditor = new Node { Type = "text", Text = "Acme", Marks = [Mark.Of("textStyle", "color", "#404040"), Mark.Of("bold")] };

        Assert.Equal(doLeitor.Fingerprint(), doEditor.Fingerprint());
    }

    [Fact]
    public void TextoVizinhoComAsMesmasMarcasEUmTextoSo()
    {
        // The reader emits one node per `w:r`: "Acme® Software" arrives split.
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
        // Merging without looking at marks would hide a formatting edit.
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
        // Loosening too much preserves the XML of a changed block, which is data loss.
        var antes = Node.Of("paragraph", new Node { Type = "text", Text = "Sumário" });
        var depois = Node.Of("paragraph", new Node { Type = "text", Text = "Sumario" });

        Assert.NotEqual(antes.Fingerprint(), depois.Fingerprint());
    }

    [Fact]
    public void OidNaoEntraNaComparacao()
    {
        // Identity is not content.
        var primeiro = Node.Of("paragraph", new Node { Type = "text", Text = "Texto" }).With("oid", "b1");
        var segundo = Node.Of("paragraph", new Node { Type = "text", Text = "Texto" }).With("oid", "b9");

        Assert.Equal(primeiro.Fingerprint(), segundo.Fingerprint());
    }

    [Fact]
    public void CaixaDeTextoViraObjetoComOTextoDentro()
    {
        // A box is its own text flow: in the line, title and subtitle would run together.
        var floats = FloatsOf(Open(Fixtures.WithTextBoxes()).Doc.Content![0]);

        Assert.Equal(2, floats.Count);
        Assert.All(floats, item => Assert.Equal("text", item.GetProperty("kind").GetString()));
        Assert.Contains(floats, item => TextOfFloat(item) == "Título do manual");
        Assert.Contains(floats, item => TextOfFloat(item) == "Subtítulo do manual");
    }

    [Fact]
    public void CaixaDeTextoNaoEntraDuasVezes()
    {
        // DrawingML and VML fallback: one branch only, or two objects in the same place.
        var floats = FloatsOf(Open(Fixtures.WithTextBoxes()).Doc.Content![0]);

        Assert.Single(floats.Where(item => TextOfFloat(item) == "Título do manual"));
    }

    [Fact]
    public void CaixaDeTextoSobreviveAEdicaoDoParagrafoAncora()
    {
        // The writer does not generate shapes from scratch: it copies anchored objects from the
        // block's original XML.
        var original = Fixtures.WithTextBoxes();
        var model = Clone(Open(original));

        // Editing the text, not deleting the `oid`, which is the thread back to the original XML.
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
        // The cover title and subtitle live in boxes: the new text has to go back.
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

        // The fallback branch too, or the file says two things.
        Assert.DoesNotContain("Título do manual", MainDocumentXml(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CaixaNaoEditadaNaoEReescrita()
    {
        // A box nobody touched keeps the XML it had.
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
        // Editable, because editing the anchor copies the shape. The warning stays: without a
        // declaration, the shape inherits frame and fill from the theme.
        var result = DocxReader.Read(Fixtures.WithTextBoxes());

        Assert.Contains(Inventory.Shapes, result.Inventory.Invisible);
        Assert.DoesNotContain(Inventory.Shapes, result.Inventory.Structural);
    }

    [Fact]
    public void CaixaSemMolduraNaoAvisaMolduraNenhuma()
    {
        // A box without decoration needs no warning.
        var result = DocxReader.Read(Fixtures.WithDecoratedTextBoxes());

        Assert.DoesNotContain(Inventory.Shapes, result.Inventory.Invisible);

        var floats = result.Model.Doc.Content!.SelectMany(FloatsOf).ToList();

        Assert.Null(TextOf(floats[0], "fill"));
        Assert.Null(TextOf(floats[0], "line"));

        // 12700 EMU is 1 pt.
        Assert.Equal("#ffffff", TextOf(floats[1], "fill"));
        Assert.Equal("#1f5fa9", TextOf(floats[1], "line"));
        Assert.Equal(1, floats[1].GetProperty("lineWidthPt").GetDouble());
        Assert.False(floats[1].GetProperty("dash").GetBoolean());
    }

    [Fact]
    public void OGradienteContinuaSendoAvisado()
    {
        // The warning now only covers what is left: gradient, texture, shadow, rounded corner.
        var result = DocxReader.Read(Fixtures.WithGradientTextBox());

        Assert.Contains(Inventory.Shapes, result.Inventory.Invisible);
    }

    [Fact]
    public void OCabecalhoTrazAFonteQueODocumentoPede()
    {
        // Without it the header would inherit the editor font.
        var floats = Open(Fixtures.WithHeaderGroup()).Page.Header!.Floats;
        var titulo = floats.Single(item => item.Kind == "text");

        var marca = titulo.Content!.Single().Content!.Single().Marks!
            .Single(m => m.Type == "textStyle");

        // No stack: the fixture does not declare the family, and the editor's `@font-face` serves
        // Calibri.
        Assert.Equal("Calibri", marca.Attrs!["fontFamily"]!.GetValue<string>());
    }

    [Fact]
    public void OFileteDoCabecalhoEDesenhado()
    {
        // The rule is a zero-height outlined shape, in the logo group.
        var floats = Open(Fixtures.WithHeaderGroup()).Page.Header!.Floats;

        var filete = Assert.Single(floats.Where(item => item.Kind == "rule"));
        Assert.True(filete.WidthMm > 100);
        Assert.Equal(0, filete.HeightMm);
    }

    [Fact]
    public void AListaTrazAMarcaEOsRecuosQueODocumentoPede()
    {
        // The document states the character and both distances, where the text starts and how far
        // before it the marker sits; the private use area character comes translated.
        var model = Open(Fixtures.WithBulletList());
        var lista = model.Doc.Content!.Single(node => node.Type == "bulletList");

        Assert.Equal("▪", lista.Attrs!["marker"]!.GetValue<string>());
        Assert.Equal(12.7, lista.Attrs["indentMm"]!.GetValue<double>(), 1);
        Assert.Equal(6.35, lista.Attrs["hangingMm"]!.GetValue<double>(), 2);
    }

    [Fact]
    public void CadaParagrafoDoRodapeAbreUmaLinha()
    {
        // Each band paragraph is a line.
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
        // The address lets saving write into the piece's `w:t` without regenerating the band.
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
        // No `w:t`, no address: typing there would erase the field.
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
        // Only the edited piece's `w:t` changes.
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
        // A part nobody touched goes back byte for byte, not reserialized by the SDK.
        var original = Fixtures.WithFooterOfThreeLines();
        var model = Clone(Open(original));

        Assert.Equal(
            PartsOf(original)["word/footer1.xml"],
            PartsOf(Save(original, model).Bytes)["word/footer1.xml"]);
    }

    [Fact]
    public void OCampoDoRodapeSobreviveAEdicaoDaLinha()
    {
        // Text merged with the tab lost its trace: saving writes nothing, before the field becomes
        // a fixed number.
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
        // The corporate header title lives in a box inside a shape group.
        var original = Fixtures.WithHeaderGroup();
        var model = Clone(Open(original));

        var caixa = Assert.Single(model.Page.Header!.Floats!, item => item.Kind == "text");
        Assert.NotNull(caixa.BoxId);

        caixa.Content![0].Content![0].Text = "EVIDÊNCIAS DE HOMOLOGAÇÃO";

        var reopened = Open(Save(original, model).Bytes).Page.Header!;
        var voltou = Assert.Single(reopened.Floats!, item => item.Kind == "text");

        Assert.Equal("EVIDÊNCIAS DE HOMOLOGAÇÃO", voltou.Content![0].Content![0].Text);

        // The group stays whole: the writer does not generate logo and rule.
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
        // `w:spacing` attributes are independent: declaring only `w:after` does not erase the
        // style's line spacing.
        var blocks = OpenFlat(Fixtures.WithStyleSpacingAndDirectMargins()).Doc.Content!;

        // 276/240 in 10 pt Arial: 1.15 × the font's natural height.
        Assert.Equal("1.3224", blocks[0].Attrs!["lineHeight"]!.GetValue<string>());
        Assert.Equal(0d, blocks[0].Attrs!["spaceAfter"]!.GetValue<double>());
    }

    [Fact]
    public void ORecuoDoParagrafoEAMedidaQueODocumentoPede()
    {
        // The editor level is worth 2.5em, 25 pt at 10 pt, not the 36 pt of 720 twips.
        var blocks = OpenFlat(Fixtures.WithStyleSpacingAndDirectMargins()).Doc.Content!;

        // 720 twips = 12.7 mm; hanging comes negative, because it is `text-indent`.
        Assert.Equal(12.7, blocks[0].Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(1.06, blocks[0].Attrs!["indentRightMm"]!.GetValue<double>());
        Assert.Equal(-6.35, blocks[0].Attrs!["firstLineMm"]!.GetValue<double>());
        Assert.Equal(0, blocks[0].Attrs!["indent"]!.GetValue<int>());

        // Changing the left indent does not erase the hanging indent from the style.
        Assert.Equal(25.4, blocks[1].Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(-6.35, blocks[1].Attrs!["firstLineMm"]!.GetValue<double>());
    }

    [Fact]
    public void ORecuoVoltaParaOArquivoNaMedidaEmQueVeio()
    {
        var original = Fixtures.WithStyleSpacingAndDirectMargins();
        var model = Clone(Open(original));

        // The edited block is rewritten from scratch.
        model.Doc.Content![0].Content![0].Text = "Outro texto no mesmo recuo.";

        // Reread flattened, the indent inherited from the style still applies.
        var reopened = OpenFlat(Save(original, model).Bytes).Doc.Content![0];
        Assert.Equal(12.7, reopened.Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(-6.35, reopened.Attrs!["firstLineMm"]!.GetValue<double>());
    }

    [Fact]
    public void OTrechoLevaSoOQueDifereDoEstilo()
    {
        // A run without its own formatting does not take the style's, or modifying the style would
        // not change the screen.
        var banner = FirstOfType(Open(Fixtures.WithStyles()), "paragraph");
        Assert.Empty(MarksOf(banner));

        // A leitura achatada, do rascunho antigo, leva tudo.
        var flat = FirstOfType(OpenFlat(Fixtures.WithStyles()), "paragraph");
        Assert.Contains(MarksOf(flat), m => m.Type == "bold");
    }

    [Fact]
    public void OBlocoLevaSoOQueOParagrafoDeclara()
    {
        // The block carries only direct formatting, at its effective value; the inherited part
        // comes from the style CSS.
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
        // The writer looks up the style's font and indent in the style.
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
        // A draft from before `.sdoc` 4 carries flattened blocks.
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
        // The section mark is an empty paragraph with `w:sectPr`, and LibreOffice gives it no
        // height.
        var original = Fixtures.WithSectionMarkInTheMiddle();
        var model = Clone(Open(original));

        var mark = model.Doc.Content![1];
        Assert.True(mark.Attrs!["sectionMark"]!.GetValue<bool>());

        // And it stays in the file, because it carries the section.
        var saved = Save(original, model);
        Assert.Equal(3, saved.Result.PreservedBlocks);
        Assert.True(Open(saved.Bytes).Doc.Content![1].Attrs!["sectionMark"]!.GetValue<bool>());
    }

    [Fact]
    public void OParagrafoQueAncoraEQuebraContinuaParagrafo()
    {
        // The cover mark anchored to the break paragraph stays on the upper sheet: the break node
        // lives between sheets.
        var model = Open(Fixtures.WithBreakOnAnchorParagraph());

        var capa = model.Doc.Content![0];
        Assert.Equal("paragraph", capa.Type);
        Assert.True(capa.Attrs!["breakAfter"]!.GetValue<bool>());
        Assert.Single(FloatsOf(capa));
    }

    [Fact]
    public void QuebraSozinhaContinuaSendoNoDeQuebra()
    {
        // With nothing anchored, a paragraph with only the break is the break node.
        var model = Open(Fixtures.WithLonePageBreak());

        Assert.Contains(model.Doc.Content!, node => node.Type == "pageBreak");
    }

    [Fact]
    public void OGiroViajaEmGrausSemMexerNasMedidas()
    {
        // Word positions unrotated and rotates around the center, like `transform: rotate()`.
        var model = Open(Fixtures.WithRotatedImage());

        var image = Assert.Single(FloatsOf(model.Doc.Content![0]));
        Assert.Equal(270, image.GetProperty("rotation").GetDouble());
        Assert.Equal(285.76, image.GetProperty("widthMm").GetDouble(), 2);
        Assert.Equal(80.14, image.GetProperty("heightMm").GetDouble(), 2);
    }

    [Fact]
    public void AAncoraViajaComOrigemEDeslocamento()
    {
        // The most common origin is the paragraph, which only has a position after pagination.
        var image = Assert.Single(FloatsOf(Open(Fixtures.WithRotatedImage()).Doc.Content![0]));

        Assert.Equal("column", image.GetProperty("hFrom").GetString());
        Assert.Equal("paragraph", image.GetProperty("vFrom").GetString());
    }

    [Theory]
    [InlineData("não é um zip")]
    [InlineData("PK zip truncado")]
    public void RejectsGarbageWithAReadableMessage(string garbage)
    {
        // A document is untrusted data: it becomes a sentence, and never brings the process down.
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

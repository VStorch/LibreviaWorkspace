using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Librevia.Format.Docx;
using Librevia.Format.Odt;
using Librevia.Format.Protocol;

namespace Librevia.Format.Tests;

/// <summary>
/// <c>odt.create</c> + <c>odt.save</c>, the only ODT generator: a document without an ODT original
/// goes into the minimal package and comes out whole, with everything the model carries.
/// </summary>
public partial class OdtCreateTests
{
    private const string Generator = "Librevia 9.9";
    private const string MimeType = "application/vnd.oasis.opendocument.text";

    private static DocumentModelDto ModelOf(string json) => JsonSerializer.Deserialize<DocumentModelDto>(json, JsonOptions.Default)!;

    private static DocumentModelDto RichModel() => ModelOf(OdtCreateFixtures.RichModel);

    /// <summary>As main does it: the minimal package, then the save into it.</summary>
    private static (byte[] Bytes, SaveResult Result) Export(DocumentModelDto model)
    {
        var template = OdtTemplate.Create(new OdtCreateDto(
            model.Page, model.Sections, model.Styles, model.Notes, Generator, model.Properties, [.. HeadingLevelsOf(model.Doc)]));
        return OdtWriter.Write(template, model);
    }

    private static IEnumerable<int> HeadingLevelsOf(Node node) =>
        (node.Type == "heading" && node.Attrs?["level"]?.GetValue<int>() is { } level ? [level] : Enumerable.Empty<int>())
        .Concat((node.Content ?? []).SelectMany(HeadingLevelsOf));

    private static List<string> Losses(SaveResult result) => [.. result.Inventory.Lost.Select(notice => notice.Key)];

    private static readonly byte[] Rich = Export(RichModel()).Bytes;

    private sealed record Entry(string Name, byte[] Data, bool Stored);

    private static List<Entry> Entries(byte[] zip)
    {
        using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        return [.. archive.Entries.Select(entry =>
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return new Entry(entry.FullName, copy.ToArray(), entry.CompressedLength == entry.Length);
        })];
    }

    private static string Part(byte[] zip, string name) => Encoding.UTF8.GetString(Entries(zip).Single(entry => entry.Name == name).Data);

    private static readonly string Content = Part(Rich, "content.xml");
    private static readonly string Styles = Part(Rich, "styles.xml");
    private static readonly string Manifest = Part(Rich, "META-INF/manifest.xml");

    private static Node Paragraph(params Node[] content) => Node.Of("paragraph", content);

    private static string TextOf(Node node) => node.Text ?? string.Concat((node.Content ?? []).Select(TextOf));

    private static Node Text(string text) => new() { Type = "text", Text = text };

    private static Node Math(string mathMl, string? latex = null, bool display = false)
    {
        var node = new Node { Type = "math" }.With("mathml", mathMl).With("display", display);
        return latex is null ? node : node.With("latex", latex);
    }

    private static DocumentModelDto WithBlocks(params Node[] blocks)
    {
        var model = RichModel();
        model.Doc.Content = [.. blocks];
        return model with { Sections = null };
    }

    [Fact]
    public void TheMimetypeOpensThePackageFirstAndStoredWithoutExtraField()
    {
        const int methodOffset = 8;
        const int nameLengthOffset = 26;
        const int extraLengthOffset = 28;
        const int headerLength = 30;
        Assert.Equal(0x04034b50, BitConverter.ToInt32(Rich, 0));
        Assert.Equal(0, BitConverter.ToInt16(Rich, methodOffset));
        Assert.Equal(0, BitConverter.ToInt16(Rich, extraLengthOffset));
        var nameLength = BitConverter.ToInt16(Rich, nameLengthOffset);
        Assert.Equal("mimetype", Encoding.ASCII.GetString(Rich, headerLength, nameLength));
        Assert.Equal(MimeType, Encoding.ASCII.GetString(Rich, headerLength + nameLength, MimeType.Length));
    }

    [Fact]
    public void TheManifestDeclaresEachPartAndEachImage()
    {
        Assert.Equal(
            ["mimetype", "content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml", "Pictures/Librevia1.png"],
            Entries(Rich).Select(entry => entry.Name));
        Assert.Contains($"manifest:full-path=\"/\" manifest:media-type=\"{MimeType}\" manifest:version=\"1.3\"", Manifest);
        Assert.Contains("manifest:full-path=\"Pictures/Librevia1.png\" manifest:media-type=\"image/png\"", Manifest);
        Assert.Contains("manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"", Manifest);
    }

    [Fact]
    public void TextAttributesAndMetadataAreEscaped()
    {
        Assert.Contains("Título &lt;script&gt;&amp;", Content);
        Assert.DoesNotContain("<script>", Content);
        Assert.Contains("xlink:href=\"https://example.com/?a=1&amp;b=2\"", Content);
        Assert.DoesNotContain("javascript:", Content);
        Assert.Contains("<dc:creator>Bia &lt;revisora&gt;</dc:creator>", Content);

        var meta = Part(Rich, "meta.xml");
        Assert.Contains($"<meta:generator>{Generator}</meta:generator>", meta);
        Assert.Contains("<dc:title>Relatório &lt;anual&gt; &amp; \"final\"</dc:title>", meta);
        Assert.Contains("<meta:keyword>contas</meta:keyword><meta:keyword>balanço</meta:keyword>", meta);
        Assert.Contains("<meta:initial-creator>Ana</meta:initial-creator>", meta);
        Assert.Contains("<meta:creation-date>2026-01-02T03:04:05Z</meta:creation-date>", meta);
    }

    [Fact]
    public void HeadingsBookmarksLinksNotesFieldsAndCommentsAreCarried()
    {
        Assert.Matches("<text:h text:style-name=\"[^\"]+\" text:outline-level=\"1\"><text:bookmark-start text:name=\"_Toc1\"/>", Content);
        Assert.Contains("<text:bookmark-end text:name=\"_Toc1\"/>", Content);
        Assert.Contains("<text:a xlink:type=\"simple\" xlink:href=\"#_Toc1\">", Content);
        Assert.Matches(
            "<text:note text:id=\"[^\"]+\" text:note-class=\"footnote\"><text:note-citation>1</text:note-citation><text:note-body><text:p[^>]*>Nota de rodapé</text:p>",
            Content);
        Assert.Contains("text:note-class=\"endnote\"><text:note-citation>i</text:note-citation>", Content);
        Assert.Contains("<text:page-number text:select-page=\"current\">1</text:page-number>", Content);
        Assert.Contains("<office:annotation office:name=\"__Annotation__9\">", Content);
        Assert.Contains("<office:annotation-end office:name=\"__Annotation__9\"/>", Content);
        Assert.Contains("<text:p>Feito</text:p>", Content);
        Assert.DoesNotContain("removido", Content);
        Assert.Contains("<office:annotation-end office:name=\"__Annotation__9\"/> inserido", Content);
    }

    [Fact]
    public void ConsecutiveSpacesTheTabAndOnlyXmlCharactersAreKept()
    {
        Assert.Contains("Olá <text:s/>", Content);

        var (bytes, _) = Export(WithBlocks(Paragraph(Text("  a\tb   c\u0001<"))));

        Assert.Contains("><text:s text:c=\"2\"/>a<text:tab/>b <text:s text:c=\"2\"/>c&lt;</text:p>", Part(bytes, "content.xml"));
    }

    [Fact]
    public void RunsGetAnAutomaticStyleAndParagraphsTheirDirectFormatting()
    {
        Assert.Matches("<style:style style:name=\"Heading_20_1\" style:family=\"paragraph\"", Styles);
        Assert.Matches("<style:style style:name=\"LvT\\d+\" style:family=\"text\"><style:text-properties fo:font-weight=\"bold\"", Content);
        Assert.Contains("fo:color=\"#c00000\"", Content);
        Assert.Contains("fo:font-family=\"Arial\"", Content);
        Assert.Contains("fo:background-color=\"#ffff00\"", Content);
        Assert.Contains("style:text-underline-style=\"solid\"", Content);
        Assert.Contains("style:text-line-through-style=\"solid\"", Content);
        Assert.Matches("fo:text-align=\"center\"[^>]*fo:margin-top=\"6pt\" fo:margin-bottom=\"12pt\"", Content);
        Assert.Contains("fo:keep-with-next=\"always\"", Content);
    }

    [Fact]
    public void ListsAreNumberedAsWordCountsThemNestedWithAStartAndContinued()
    {
        Assert.Matches("<text:list text:style-name=\"LvL1\"><text:list-item>", Content);
        Assert.Matches("<text:list-level-style-number text:level=\"1\"[^>]* style:num-suffix=\".\" style:num-format=\"1\" text:start-value=\"3\">", Content);
        Assert.Matches("text:level=\"2\"[^>]* style:num-suffix=\"\\)\" style:num-format=\"a\" text:display-levels=\"2\"", Content);
        Assert.Contains("<text:list-level-style-bullet text:level=\"1\" text:bullet-char=\"•\">", Content);

        var list = RichModel().Doc.Content![3];
        var (bytes, _) = Export(WithBlocks(list, Paragraph(Text("entre")), list));

        Assert.Contains("<text:list text:style-name=\"LvL1\" text:continue-numbering=\"true\">", Part(bytes, "content.xml"));
    }

    [Fact]
    public void ATableInAListItemGoesRightAfterTheList()
    {
        var table = RichModel().Doc.Content![6];
        var item = Node.Of("listItem", Paragraph(Text("item")), table);
        var (bytes, _) = Export(WithBlocks(Node.Of("bulletList", item)));

        Assert.Matches("</text:list><table:table ", Part(bytes, "content.xml"));
    }

    [Fact]
    public void MergedCellsGetCoveredCellsAndTheImageGoesInThePackage()
    {
        Assert.Contains("table:number-columns-spanned=\"2\"", Content);
        Assert.Contains("table:number-rows-spanned=\"2\"", Content);
        Assert.Matches("<table:table-cell[^>]*><text:p[^>]*>C<text:note [^>]*>.*Nota na célula", Content);
        Assert.Equal(2, Regex.Count(Content, "<table:covered-table-cell/>"));
        Assert.Contains("fo:border-top=\"1.5pt double #ff0000\"", Content);
        Assert.Contains("<draw:image xlink:href=\"Pictures/Librevia1.png\"", Content);
        Assert.Contains("svg:width=\"25.4mm\" svg:height=\"12.7mm\"", Content);
    }

    [Fact]
    public void EachSectionGetsItsMasterPageAndTheBreakGoesToTheNextBlock()
    {
        Assert.Matches("style:master-page-name=\"Section1\"><style:paragraph-properties style:page-number=\"1\"/>", Content);
        Assert.Matches("style:master-page-name=\"Standard\"><style:paragraph-properties style:page-number=\"auto\"/>", Content);
        Assert.Contains("fo:break-before=\"page\"", Content);
        Assert.Contains("<style:master-page style:name=\"Section1\" style:page-layout-name=\"LvPm1\"", Styles);
        Assert.Contains("<style:master-page style:name=\"Standard\" style:page-layout-name=\"LvPm2\">", Styles);
        Assert.Contains("fo:page-width=\"210mm\" fo:page-height=\"297mm\"", Styles);
        Assert.Contains("fo:margin-left=\"30mm\"", Styles);

        var header = HeaderOf().Match(Styles).Groups[1].Value;
        Assert.Contains("Relatório — página <text:page-number text:select-page=\"current\">1</text:page-number>", header);
        Assert.Contains("<text:page-count>1</text:page-count>", header);
        var single = Export(WithBlocks(Paragraph(Text("Só a página")))).Bytes;
        Assert.Equal("Relatório — página {n} de {total}", OdtReader.Read(single).Model.Page.HeaderText);
    }

    [Fact]
    public void ColumnsGoInATextSection()
    {
        var model = RichModel();
        model = model with { Page = model.Page with { Columns = new ColumnsDto(2, 10, Separator: true) } };

        var content = Part(Export(model).Bytes, "content.xml");

        Assert.Matches("<text:section text:style-name=\"LvSect1\" text:name=\"LvSection1\"><text:p[^>]*>Segunda seção</text:p></text:section>", content);
        Assert.Matches("<style:columns fo:column-count=\"2\" fo:column-gap=\"10mm\"><style:column-sep ", content);
    }

    [Fact]
    public void OnlyTheXmlPartsAreCompressed()
    {
        var compressed = Entries(Rich).Where(entry => !entry.Stored).Select(entry => entry.Name);

        Assert.Equal(["content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml"], compressed);
    }

    [Fact]
    public void EachEquationIsAFormulaObjectLikeACharacterWithItsMathMlInItsFolder()
    {
        var (bytes, result) = Export(WithBlocks(
            Paragraph(Text("Seja "), Math("<math display=\"inline\"><msup><mi>x</mi><mn>2</mn></msup></math>", "x^2"), Text(" & mais")),
            Paragraph(Math("<math display=\"block\"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>", display: true)),
            Paragraph(Math("<script/>", "y"))));
        var content = Part(bytes, "content.xml");

        Assert.Matches(
            "<draw:frame draw:style-name=\"LvFr\\d+\" draw:name=\"LvEquation\\d+\" text:anchor-type=\"as-char\" svg:width=\"[\\d.]+mm\" svg:height=\"[\\d.]+mm\" draw:z-index=\"0\"><draw:object xlink:href=\"\\./Object 1\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/><svg:desc>x\\^2</svg:desc></draw:frame>",
            content);
        Assert.Contains("xlink:href=\"./Object 2\"", content);
        Assert.DoesNotContain("Object 3", content);
        Assert.DoesNotContain("script", content);
        Assert.Contains(new Notice(Inventory.Equations), result.Inventory.Lost);
        Assert.Equal(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><msup><mi>x</mi><mn>2</mn></msup></math>",
            Part(bytes, "Object 1/content.xml"));
        Assert.Contains("display=\"block\"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>", Part(bytes, "Object 2/content.xml"));
    }

    [Theory]
    [InlineData("<!DOCTYPE math [<!ENTITY x \"y\">]><math><mi>&x;</mi></math>")]
    [InlineData("<math><!-- x --><mi>x</mi></math>")]
    [InlineData("<math><mi>&nbsp;</mi></math>")]
    [InlineData("<math xmlns=\"http://www.w3.org/1999/xhtml\"><mi>x</mi></math>")]
    [InlineData("<m:math xmlns:m=\"http://www.w3.org/1998/Math/MathML\"><m:mi>x</m:mi></m:math>")]
    [InlineData("<mrow><mi>x</mi></mrow>")]
    public void MathMlThatIsNotPlainMathMlLeavesNoObject(string mathMl)
    {
        var (bytes, _) = Export(WithBlocks(Paragraph(Math(mathMl))));

        Assert.DoesNotContain("draw:object", Part(bytes, "content.xml"));
    }

    [Fact]
    public void OnlyAllowedElementsAndAttributesReachTheFormula()
    {
        var (bytes, _) = Export(WithBlocks(Paragraph(Math(
            "<math><mi href=\"javascript:x\" mathvariant=\"bold\" class=\"omml-box evil\">x</mi><maction><mi>y</mi></maction></math>"))));

        Assert.EndsWith("<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi mathvariant=\"bold\" class=\"omml-box\">x</mi></math>", Part(bytes, "Object 1/content.xml"));
    }

    [Fact]
    public void TheManifestDeclaresEachObjectAndEachOfItsPartsAndOnlyWhatIsInThePackage()
    {
        var (bytes, _) = Export(WithBlocks(
            Paragraph(Math("<math><mi>x</mi></math>")),
            Paragraph(Math("<math><mi>y</mi></math>"))));
        var manifest = Part(bytes, "META-INF/manifest.xml");

        Assert.Contains(
            "<manifest:file-entry manifest:full-path=\"Object 1/\" manifest:version=\"1.3\" manifest:media-type=\"application/vnd.oasis.opendocument.formula\"/>",
            manifest);
        Assert.Contains("manifest:full-path=\"Object 2/content.xml\" manifest:media-type=\"text/xml\"", manifest);
        var declared = FullPath().Matches(manifest).Select(match => match.Groups[1].Value).ToList();
        var files = declared.Where(path => path != "/" && !path.EndsWith('/')).Order(StringComparer.Ordinal);
        var entries = Entries(bytes).Select(entry => entry.Name).Where(name => name != "mimetype" && !name.StartsWith("META-INF/", StringComparison.Ordinal));
        Assert.Equal(entries.Order(StringComparer.Ordinal), files);
    }

    [Fact]
    public void TheFrameGrowsWithTheTextAndWithWhatTheEquationStacks()
    {
        var (bytes, _) = Export(WithBlocks(
            Paragraph(Math("<math><mi>x</mi></math>")),
            Paragraph(Math("<math><mfrac><mi>x</mi><mn>2</mn></mfrac></math>"))));
        var sizes = FrameSize().Matches(Part(bytes, "content.xml"))
            .Select(match => (Width: double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture),
                Height: double.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture)))
            .ToList();

        Assert.True(sizes[1].Height > sizes[0].Height);
        Assert.True(sizes[1].Width > sizes[0].Width);
    }

    [Fact]
    public void AnAnchoredImageIsAFrameOnTheParagraph()
    {
        var floats = new JsonArray(new JsonObject
        {
            ["kind"] = "image",
            ["src"] = RichModel().Doc.Content![2].Content!.Last().Attrs!["src"]!.GetValue<string>(),
            ["widthMm"] = 20, ["heightMm"] = 30, ["rotation"] = 0,
            ["hFrom"] = "margin", ["hAlign"] = "right", ["vFrom"] = "paragraph", ["vOffsetMm"] = 2,
            ["behind"] = false, ["wrap"] = "square",
        });
        var (bytes, _) = Export(WithBlocks(Paragraph(Text("ao lado")).With("floats", floats)));
        var content = Part(bytes, "content.xml");

        Assert.Matches("<text:p[^>]*><draw:frame draw:style-name=\"LvFr1\" draw:name=\"LvFrame1\" text:anchor-type=\"paragraph\" svg:y=\"2mm\" svg:width=\"20mm\" svg:height=\"30mm\"", content);
        Assert.Contains("style:horizontal-pos=\"right\" style:horizontal-rel=\"page-content\"", content);
    }

    [Fact]
    public void AHeaderFromAnotherFormatKeepsItsColumnsAndRule()
    {
        var model = RichModel();
        var band = new BandDto(
            [new PieceDto("text", "Empresa", Bold: true)],
            [new PieceDto("pageNumber")],
            [new PieceDto("text", "{n}", Literal: true)],
            Rule: true);
        model = model with { Sections = null, Page = model.Page with { Header = band, HeaderText = null } };

        var (bytes, _) = Export(model);
        var styles = Part(bytes, "styles.xml");

        var header = HeaderOf().Match(styles).Groups[1].Value;
        Assert.Matches("^<text:p text:style-name=\"LvMP\\d+\"><text:span text:style-name=\"LvT\\d+\">Empresa</text:span><text:tab ?/><text:page-number text:select-page=\"current\">1</text:page-number><text:tab ?/>\\{n\\}</text:p>$", header);
        Assert.Contains("<style:tab-stop style:type=\"center\" style:position=\"80mm\"", styles);
        Assert.Contains("fo:border-bottom=\"0.5pt solid #000000\"", styles);
    }

    [Fact]
    public void LineSpacingWithoutItsOwnFontReadsBackAsTheSameModel()
    {
        var spaced = Paragraph(Text("Entrelinha")).With("lineHeight", "1.5");

        var (bytes, _) = Export(WithBlocks(spaced));

        var read = OdtReader.Read(bytes).Model.Doc.Content!.Single();
        Assert.Equal("1.5", read.Attrs!["lineHeight"]!.ToString());
        Assert.False(read.Attrs.ContainsKey("fontFamily"));
    }

    [Fact]
    public void TheCreatedPackageReadsBackWithEveryBlockAndItsStyles()
    {
        var read = OdtReader.Read(Rich).Model;

        Assert.Equal("Segunda seção", TextOf(read.Doc.Content!.Last()));
        Assert.Equal("heading 1", read.Styles!.Styles["Heading1"].Name);
        Assert.Equal(RichModel().Doc.Content![0].Attrs!["level"]!.GetValue<int>(), read.Doc.Content![0].Attrs!["level"]!.GetValue<int>());
    }

    [Fact]
    public void SavingTheSameModelGivesTheSameBytes()
    {
        var template = OdtTemplate.Create(new OdtCreateDto(RichModel().Page, Styles: RichModel().Styles, Generator: Generator));

        Assert.Equal(template, OdtTemplate.Create(new OdtCreateDto(RichModel().Page, Styles: RichModel().Styles, Generator: Generator)));
    }

    [Fact]
    public void ACodeBlockIsOneMonospacedParagraphPerLineWithItsSpacesKept()
    {
        var (bytes, result) = Export(WithBlocks(Node.Of("codeBlock", Text("if (a)\n    b();"))));

        var content = Part(bytes, "content.xml");
        var style = Regex.Match(content, "<text:p text:style-name=\"(LvP\\d+)\">if \\(a\\)</text:p><text:p text:style-name=\"\\1\"><text:s text:c=\"4\"/>b\\(\\);</text:p>").Groups[1].Value;
        Assert.NotEmpty(style);
        Assert.Matches($"style:name=\"{style}\"[^>]*><style:paragraph-properties fo:margin-top=\"0pt\" fo:margin-bottom=\"0pt\"/><style:text-properties fo:font-family=\"'Liberation Mono'\"", content);
        Assert.Empty(Losses(result));
    }

    [Fact]
    public void AHorizontalRuleIsAnEmptyParagraphWithABottomBorder()
    {
        var (bytes, result) = Export(WithBlocks(Paragraph(Text("antes")), Node.Of("horizontalRule"), Paragraph(Text("depois"))));

        var content = Part(bytes, "content.xml");
        var style = Regex.Match(content, "<text:p text:style-name=\"(LvP\\d+)\"/>").Groups[1].Value;
        Assert.Matches($"style:name=\"{style}\"[^>]*><style:paragraph-properties fo:border-bottom=\"0.5pt solid #000000\"/>", content);
        Assert.Empty(Losses(result));
    }

    [Fact]
    public void AHeadingWithoutItsStyleInTheSheetGetsTheBuiltInHeadingOfItsLevel()
    {
        var model = WithBlocks(Node.Of("heading", Text("Seção")).With("level", 2));
        var sheet = model.Styles!;
        var styles = sheet.Styles.Where(entry => entry.Key != "Heading2").ToDictionary();

        var (bytes, _) = Export(model with { Styles = sheet with { Styles = styles } });

        Assert.Matches("<style:style style:name=\"Heading_20_2\" style:family=\"paragraph\"[^>]* style:default-outline-level=\"2\">" +
                       "<style:paragraph-properties fo:margin-top=\"17pt\" fo:margin-bottom=\"14.1pt\" fo:keep-with-next=\"always\" ?/>" +
                       "<style:text-properties fo:font-size=\"17pt\" fo:font-weight=\"bold\" ?/>", Part(bytes, "styles.xml"));
        Assert.Matches("<text:h text:style-name=\"Heading_20_2\" text:outline-level=\"2\">Seção</text:h>", Part(bytes, "content.xml"));
    }

    [Fact]
    public void AHeadingTakesTheSheetStyleNamedAfterItsLevelWhateverItsId()
    {
        var model = WithBlocks(Node.Of("heading", Text("Seção")).With("level", 3));
        var sheet = model.Styles!;
        var styles = sheet.Styles.Where(entry => entry.Key != "Heading3").ToDictionary();
        styles["Ttulo3"] = sheet.Styles["Heading3"] with { Id = "Ttulo3" };

        var (bytes, _) = Export(model with { Styles = sheet with { Styles = styles } });

        Assert.DoesNotContain("Heading_20_3", Part(bytes, "styles.xml"));
        Assert.Contains("<text:h text:style-name=\"Ttulo3\" text:outline-level=\"3\">", Part(bytes, "content.xml"));
    }

    [Fact]
    public void ANewFileNamesTheCreatorAsTheLastAuthorAndStatesTheEditingTime()
    {
        var model = RichModel();
        var (bytes, _) = Export(model with { Properties = model.Properties! with { TotalTime = 90 } });

        var meta = Part(bytes, "meta.xml");
        Assert.Contains("<dc:creator>Ana</dc:creator>", meta);
        Assert.Contains("<meta:editing-duration>PT90M</meta:editing-duration>", meta);

        var (named, _) = Export(model with { Properties = model.Properties! with { LastModifiedBy = "Bia" } });
        Assert.Contains("<dc:creator>Bia</dc:creator>", Part(named, "meta.xml"));
    }

    [Fact]
    public void SavingAnOpenedFileLeavesTheAuthorAndEditingTimeThePersonDidNotTouch()
    {
        var opened = OdtReader.Read(Rich).Model;

        var (bytes, _) = OdtWriter.Write(Rich, opened with { Properties = opened.Properties! with { TotalTime = 500 } });

        Assert.Equal(Part(Rich, "meta.xml"), Part(bytes, "meta.xml"));
    }

    [Fact]
    public void FootnotesThatRestartOnEachPageRestartOnEachPage()
    {
        var model = RichModel() with { Notes = new NotesDto(new NotePrDto(Restart: "eachPage"), new NotePrDto(Restart: "eachPage")) };

        var styles = Part(Export(model).Bytes, "styles.xml");

        Assert.Matches("text:note-class=\"footnote\"[^>]* text:start-numbering-at=\"page\"", styles);
        Assert.Matches("text:note-class=\"endnote\"[^>]* text:start-numbering-at=\"document\"", styles);
    }

    [Fact]
    public void AHangingIndentPullsTheFirstLineBackAndReadsBackAsWordsNegativeFirstLine()
    {
        var hanging = Paragraph(Text("Pendente")).With("indentMm", 10.0).With("hangingMm", 5.0);

        var (bytes, _) = Export(WithBlocks(hanging));

        Assert.Matches("fo:margin-left=\"10mm\" fo:text-indent=\"-5mm\"", Part(bytes, "content.xml"));
        var read = OdtReader.Read(bytes).Model.Doc.Content!.Single();
        Assert.Equal((10.0, -5.0), (read.Attrs!["indentMm"]!.GetValue<double>(), read.Attrs["firstLineMm"]!.GetValue<double>()));
        Assert.False(read.Attrs.ContainsKey("hangingMm"));
    }

    [GeneratedRegex("<style:header>(.*?)</style:header>")]
    private static partial Regex HeaderOf();

    [GeneratedRegex("manifest:full-path=\"([^\"]+)\"")]
    private static partial Regex FullPath();

    [GeneratedRegex("svg:width=\"([\\d.]+)mm\" svg:height=\"([\\d.]+)mm\" draw:z-index=\"0\"><draw:object")]
    private static partial Regex FrameSize();
}

using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using Librevia.Format.Docx;
using Librevia.Format.Odt;
using static Librevia.Format.Tests.OdtFixtures;

namespace Librevia.Format.Tests;

/// <summary>
/// <c>odt.open</c>: an ODT becomes the same model as a DOCX, and what the model does not carry goes
/// to the inventory as lost, never silently.
/// </summary>
public class OdtReaderTests
{
    /// <summary>Accents and symbols as they are, so the expected text reads as the document.</summary>
    private static readonly JsonSerializerOptions Readable = new(DocxJson.Options)
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>Without the <c>oid</c>s: identity, which <see cref="OdtWriterTests"/> covers.</summary>
    private static OpenResult Open(byte[] bytes)
    {
        var result = OdtReader.Read(bytes);
        foreach (var block in result.Model.Doc.Content ?? [])
        {
            block.Attrs?.Remove("oid");
            if (block.Attrs is { Count: 0 }) block.Attrs = null;
        }

        return result;
    }

    private static List<Node> Blocks(byte[] bytes) => Open(bytes).Model.Doc.Content!;

    private static Node Only(string body, string automatic = "") => Blocks(Build(body, automatic))[0];

    private static string Json(object? value) => JsonSerializer.Serialize(value, Readable);

    private static List<string> Noted(byte[] bytes) =>
        [.. Open(bytes).Inventory.Invisible.Select(notice => notice.Value is null ? notice.Key : $"{notice.Key}={notice.Value}")];

    private static string? AttrOf(Node node, string name) => node.Attrs?.GetValueOrDefault(name)?.ToJsonString(Readable);

    [Fact]
    public void LibreOfficeParagraphKeepsItsNamedStyleAndOnlyItsDirectFormatting()
    {
        var result = Open(LibreOffice());
        var blocks = result.Model.Doc.Content!;

        Assert.Equal("heading", blocks[0].Type);
        Assert.Equal("1", AttrOf(blocks[0], "level"));
        Assert.Equal("\"Heading1\"", AttrOf(blocks[0], "styleId"));

        var body = blocks[1];
        Assert.Equal("paragraph", body.Type);
        Assert.Equal(
            """{"styleId":"BodyText","textAlign":"center"}""",
            Json(body.Attrs));
        Assert.Equal(
            """[{"type":"text","text":"Olá "},{"type":"text","text":"mundo","marks":[{"type":"bold"}]},""" +
            """{"type":"text","text":" e  mais\tfim"},{"type":"hardBreak"},{"type":"text","text":"linha "},""" +
            """{"type":"text","text":"forte","marks":[{"type":"bold"},{"type":"charStyle","attrs":{"styleId":"Strong"}}]}]""",
            Json(body.Content));

        Assert.Equal("Corpo", blocks[2].Content![0].Text);
        Assert.Equal(3, blocks.Count);
        Assert.True(result.Inventory.IsEmpty);
    }

    [Fact]
    public void NamedStylesBecomeTheStyleSheetWithWordsBuiltins()
    {
        var styles = Open(LibreOffice()).Model.Styles!;

        Assert.Equal("Normal", styles.Defaults.ParagraphStyleId);
        Assert.Equal("Normal", styles.Styles["Normal"].Name);
        Assert.DoesNotContain("Standard", styles.Styles.Keys);
        Assert.Equal("Liberation Serif, serif", styles.Defaults.Character.FontFamily);
        Assert.Equal("12pt", styles.Defaults.Character.FontSize);

        var body = styles.Styles["BodyText"];
        Assert.Equal(("Body Text", "paragraph", "Normal"), (body.Name, body.Type, body.BasedOn));
        Assert.Equal(7, body.Paragraph!.SpaceAfter);
        Assert.Equal(new LineSpacingDto("multiple", Factor: 1.15), body.Paragraph.LineSpacing);

        var heading = styles.Styles["Heading1"];
        Assert.Equal(("heading 1", "Heading", "BodyText"), (heading.Name, heading.BasedOn, heading.Next));
        Assert.Equal(0, heading.Paragraph!.OutlineLevel);
        Assert.Equal("18.2pt", heading.Character!.FontSize);
        Assert.True(heading.Character.Bold);
        Assert.Equal("Liberation Sans, sans-serif", styles.Styles["Heading"].Character!.FontFamily);

        Assert.Equal(("Strong", "character"), (styles.Styles["Strong"].Name, styles.Styles["Strong"].Type));
    }

    [Theory]
    [InlineData("Title", "Title", "Title")]
    [InlineData("Subtitle", "Subtitle", "Subtitle")]
    [InlineData("Quotations", "Quote", "Quote")]
    [InlineData("List_20_Paragraph", "ListParagraph", "List Paragraph")]
    [InlineData("Heading_20_9", "Heading9", "heading 9")]
    [InlineData("Minha_20_nota", "Minha_20_nota", "Minha nota")]
    public void LibreOfficeBuiltinTakesWordsIdAndName(string odfName, string id, string name)
    {
        var result = Open(Build(
            $"""<text:p text:style-name="{odfName}">x</text:p>""",
            styles: LibreOfficeStyles + $"""<style:style style:name="{odfName}" style:family="paragraph" style:parent-style-name="Standard"/>"""));

        Assert.Equal($"\"{id}\"", AttrOf(result.Model.Doc.Content![0], "styleId"));
        Assert.Equal((name, "Normal"), (result.Model.Styles!.Styles[id].Name, result.Model.Styles.Styles[id].BasedOn));
    }

    [Fact]
    public void AutomaticTextStylesBecomeMarks()
    {
        var paragraph = Only(
            """<text:p><text:span text:style-name="T1">a</text:span><text:span text:style-name="T2">b</text:span><text:span text:style-name="T3">c</text:span><text:span text:style-name="T4">d</text:span></text:p>""",
            """<style:style style:name="T1" style:family="text"><style:text-properties fo:font-style="italic" style:text-underline-style="solid" style:text-line-through-style="solid"/></style:style>""" +
            """<style:style style:name="T2" style:family="text"><style:text-properties style:text-position="super 58%" fo:color="#C00000" fo:font-size="14pt" style:font-name="Liberation Sans"/></style:style>""" +
            """<style:style style:name="T3" style:family="text"><style:text-properties style:text-position="-33% 58%" fo:background-color="#ffff00"/></style:style>""" +
            """<style:style style:name="T4" style:family="text"><style:text-properties fo:text-transform="uppercase" fo:font-variant="small-caps"/></style:style>""");

        Assert.Equal(
            """[{"type":"text","text":"a","marks":[{"type":"italic"},{"type":"strike"},{"type":"underline"}]},""" +
            """{"type":"text","text":"b","marks":[{"type":"superscript"},{"type":"textStyle","attrs":{"color":"#c00000","fontSize":"14pt","fontFamily":"Liberation Sans, sans-serif"}}]},""" +
            """{"type":"text","text":"c","marks":[{"type":"subscript"},{"type":"highlight","attrs":{"color":"#ffff00"}}]},""" +
            """{"type":"text","text":"d","marks":[{"type":"caps"},{"type":"smallCaps"}]}]""",
            Json(paragraph.Content));
    }

    [Fact]
    public void TextTheStyleMakesBoldAndTheSpanUndoesIsMarkedOff()
    {
        var heading = Only(
            """<text:h text:style-name="Heading_20_1" text:outline-level="1">a<text:span text:style-name="T1">b</text:span></text:h>""",
            """<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="normal"/></style:style>""");

        Assert.Null(heading.Content![0].Marks);
        Assert.Equal("""[{"type":"bold","attrs":{"off":true}}]""", Json(heading.Content[1].Marks));
    }

    [Fact]
    public void ExternalLinksStayAndInternalOnesGoWithTheirBookmark()
    {
        var paragraph = Only(
            """<text:p><text:a xlink:type="simple" xlink:href="https://example.com/?a=1&amp;b=2">site</text:a> <text:a xlink:href="#Intro">dentro</text:a></text:p>""");

        Assert.Equal(
            """[{"type":"text","text":"site","marks":[{"type":"link","attrs":{"href":"https://example.com/?a=1&b=2"}}]},{"type":"text","text":" dentro"}]""",
            Json(paragraph.Content));
    }

    [Fact]
    public void DirectParagraphFormattingSitsOverTheNamedStyle()
    {
        var paragraph = Only(
            """<text:p text:style-name="P1">x</text:p>""",
            """<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Text_20_body"><style:paragraph-properties fo:margin-left="1cm" fo:margin-right="0.5cm" fo:text-indent="-0.5cm" fo:margin-top="6pt" fo:margin-bottom="0cm" fo:line-height="200%" fo:keep-with-next="always" fo:keep-together="always" fo:widows="0" fo:orphans="0" fo:background-color="#eeeeee" fo:break-after="page"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="11pt"/></style:style>""");

        Assert.Equal(
            """
            {"styleId":"BodyText","breakAfter":true,"indentMm":10,"indentRightMm":5,"firstLineMm":-5,"background":"#eeeeee","spaceBefore":6,"spaceAfter":0,"lineHeight":"2.2998","fontFamily":"Liberation Sans, sans-serif","fontSize":"11pt","keepNext":true,"keepLines":true,"widowControl":false}
            """,
            Json(paragraph.Attrs));
    }

    [Fact]
    public void ALibreOfficeHangingIndentIsTheMarginOfTheTextAndANegativeFirstLineAsDocxReadsIt()
    {
        var paragraph = Only(
            """<text:p text:style-name="P1">x</text:p>""",
            """<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-left="1.27cm" fo:margin-right="0cm" fo:text-indent="-0.635cm" style:auto-text-indent="false"/></style:style>""");

        Assert.Equal(12.7, paragraph.Attrs!["indentMm"]!.GetValue<double>());
        Assert.Equal(-6.35, paragraph.Attrs["firstLineMm"]!.GetValue<double>());
        Assert.False(paragraph.Attrs.ContainsKey("hangingMm"));
    }

    [Fact]
    public void PageBreakBeforeABlockIsAPageBreakNode()
    {
        var blocks = Blocks(Build(
            """<text:p>um</text:p><text:p text:style-name="P1">dois</text:p>""",
            """<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-before="page"/></style:style>"""));

        Assert.Equal(["paragraph", "pageBreak", "paragraph"], blocks.Select(block => block.Type));
        Assert.Null(blocks[2].Attrs);
    }

    [Fact]
    public void ListsCarryTheirStyleLevelsAndNest()
    {
        var list = Only(
            """<text:list text:style-name="L1"><text:list-item text:start-value="3"><text:p>três</text:p></text:list-item>""" +
            """<text:list-item><text:p>quatro</text:p><text:list><text:list-item><text:p>sub</text:p></text:list-item></text:list></text:list-item></text:list>""",
            """<text:list-style style:name="L1">""" +
            """<text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number>""" +
            """<text:list-level-style-bullet text:level="2" text:bullet-char="◦"><style:list-level-properties text:space-before="1.5cm" text:min-label-width="0.5cm"/></text:list-level-style-bullet>""" +
            """</text:list-style>""");

        Assert.Equal("orderedList", list.Type);
        Assert.Equal(
            """{"key":"odt1","levels":[{"fmt":"decimal","text":"%1.","start":3,"indentMm":12.7,"hangingMm":6.35},""" +
            """{"fmt":"bullet","text":"◦","start":1,"indentMm":20,"hangingMm":5}]}""",
            AttrOf(list, "numbering"));

        var items = list.Content!;
        Assert.Equal("três", items[0].Content![0].Content![0].Text);
        var nested = items[1].Content![1];
        Assert.Equal("bulletList", nested.Type);
        Assert.NotNull(AttrOf(nested, "numbering"));
        Assert.Contains("\"key\":\"odt1\"", AttrOf(nested, "numbering"));
    }

    [Fact]
    public void ListItemsCarryTheirEffectiveFormatting()
    {
        var list = Only(
            """<text:list text:style-name="L1"><text:list-item><text:p text:style-name="Text_20_body">x</text:p></text:list-item></text:list>""",
            """<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"/></text:list-style>""");

        Assert.Equal(
            """{"styleId":"BodyText","spaceBefore":0,"spaceAfter":7,"lineHeight":"1.3224","fontFamily":"Liberation Serif, serif","fontSize":"12pt"}""",
            Json(list.Content![0].Content![0].Attrs));
    }

    [Fact]
    public void ListStyleNamedByTheParagraphAndLibreOfficeNumberFormat()
    {
        var list = Only(
            """<text:list><text:list-item><text:p text:style-name="P1">x</text:p></text:list-item></text:list>""",
            """<style:style style:name="P1" style:family="paragraph" style:list-style-name="L1"/>""" +
            """<text:list-style style:name="L1"><text:list-level-style-number text:level="1" loext:num-list-format="(%1%)" style:num-prefix="(" style:num-suffix=")" style:num-format="a"/></text:list-style>""");

        Assert.Equal("""{"key":"odt1","levels":[{"fmt":"lowerLetter","text":"(%1)","start":1}]}""", AttrOf(list, "numbering"));
    }

    [Fact]
    public void ContinuedListSharesTheKeyAndAnotherRestarts()
    {
        const string Style = """<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1"/></text:list-style>""";
        const string Item = "<text:list-item><text:p>x</text:p></text:list-item>";
        var blocks = Blocks(Build(
            $"""<text:list xml:id="list1" text:style-name="L1">{Item}</text:list><text:p>meio</text:p>""" +
            $"""<text:list text:continue-numbering="true" text:style-name="L1">{Item}</text:list>""" +
            $"""<text:list text:style-name="L1">{Item}</text:list>""" +
            $"""<text:list text:continue-list="list1" text:style-name="L1">{Item}</text:list>""",
            Style));

        string KeyOf(Node list) => JsonNode.Parse(AttrOf(list, "numbering")!)!["key"]!.GetValue<string>();
        Assert.Equal(["odt1", "odt1", "odt2", "odt1"], new[] { blocks[0], blocks[2], blocks[3], blocks[4] }.Select(KeyOf));
    }

    [Fact]
    public void AnItemHoldingOnlyADeeperListGoesUnderThePreviousItem()
    {
        var list = Only(
            """<text:list><text:list-item><text:p>a</text:p></text:list-item><text:list-item><text:list><text:list-item><text:p>b</text:p></text:list-item></text:list></text:list-item></text:list>""");

        Assert.Single(list.Content!);
        Assert.Equal(["paragraph", "bulletList"], list.Content![0].Content!.Select(child => child.Type));
    }

    [Fact]
    public void UnknownNumberFormatIsDecimalWithANotice()
    {
        var bytes = Build(
            """<text:list text:style-name="L1"><text:list-item><text:p>x</text:p></text:list-item></text:list>""",
            """<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="א"/></text:list-style>""");

        Assert.Contains("\"fmt\":\"decimal\"", AttrOf(Blocks(bytes)[0], "numbering"));
        Assert.Equal(["sidecar.notice.listNumberFormat=א"], Noted(bytes));
    }

    [Fact]
    public void TablesKeepSpansWidthsHeaderRowsBordersAndShading()
    {
        var table = Only(
            """<table:table table:name="T"><table:table-column table:style-name="co1" table:number-columns-repeated="2"/><table:table-column table:style-name="co2"/>""" +
            """<table:table-header-rows><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>A</text:p></table:table-cell><table:covered-table-cell/><table:table-cell table:number-rows-spanned="2"><text:p>B</text:p></table:table-cell></table:table-row></table:table-header-rows>""" +
            """<table:table-row><table:table-cell table:style-name="ce1"><text:p>C</text:p></table:table-cell><table:table-cell table:style-name="ce2"/><table:covered-table-cell/></table:table-row></table:table>""",
            """<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2.54cm"/></style:style>""" +
            """<style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>""" +
            """<style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#EEEEEE" fo:border="0.5pt solid #000000" fo:border-top="1.5pt double #ff0000" fo:border-left="none"/></style:style>""" +
            """<style:style style:name="ce2" style:family="table-cell"><style:table-cell-properties fo:background-color="transparent" fo:border="0.05pt dotted #808080"/></style:style>""");

        Assert.Equal("table", table.Type);
        var header = table.Content![0].Content!;
        Assert.Equal(["tableHeader", "tableHeader"], header.Select(cell => cell.Type));
        Assert.Equal("""{"colspan":2,"rowspan":1,"colwidth":[96,96],"borders":"top:none,0.25,#000000;right:none,0.25,#000000;bottom:none,0.25,#000000;left:none,0.25,#000000"}""", Json(header[0].Attrs));
        Assert.Equal("""{"colspan":1,"rowspan":2,"colwidth":[96],"borders":"top:none,0.25,#000000;right:none,0.25,#000000;bottom:none,0.25,#000000;left:none,0.25,#000000"}""", Json(header[1].Attrs));

        var body = table.Content[1].Content!;
        Assert.Equal(
            """{"colspan":1,"rowspan":1,"colwidth":[96],"shading":"#eeeeee","borders":"top:double,1.5,#ff0000;left:none,0.25,#000000"}""",
            Json(body[0].Attrs));
        Assert.Equal(
            """{"colspan":1,"rowspan":1,"colwidth":[96],"borders":"top:dotted,0.25,#808080;right:dotted,0.25,#808080;bottom:dotted,0.25,#808080;left:dotted,0.25,#808080"}""",
            Json(body[1].Attrs));
        Assert.Equal("paragraph", body[1].Content![0].Type);
    }

    [Fact]
    public void PicturesBecomeImagesWithTheirSizeAndDescription()
    {
        var png = Png();
        var bytes = Build(
            """<text:p><draw:frame draw:name="Imagem1" text:anchor-type="as-char" svg:width="25.4mm" svg:height="1.27cm"><draw:image xlink:href="Pictures/a.png" xlink:type="simple"/><svg:desc>figura</svg:desc></draw:frame></text:p>""",
            files: new Dictionary<string, byte[]> { ["Pictures/a.png"] = png },
            manifestExtra: """<manifest:file-entry manifest:full-path="Pictures/a.png" manifest:media-type="image/png"/>""");

        var image = Blocks(bytes)[0].Content![0];
        Assert.Equal("image", image.Type);
        Assert.Equal(
            $$"""{"src":"data:image/png;base64,{{Convert.ToBase64String(png)}}","alt":"figura","width":96,"height":48}""",
            Json(image.Attrs));
        Assert.Empty(Noted(bytes));
    }

    [Fact]
    public void FloatingMissingAndForeignPicturesGoToTheInventory()
    {
        var bytes = Build(
            """<text:p><draw:frame text:anchor-type="paragraph"><draw:image xlink:href="Pictures/a.png"/></draw:frame>""" +
            """<draw:frame text:anchor-type="as-char"><draw:image xlink:href="Pictures/falta.png"/></draw:frame>""" +
            """<draw:frame text:anchor-type="as-char"><draw:image xlink:href="Pictures/b.svg"/></draw:frame>""" +
            """<draw:frame text:anchor-type="as-char"><draw:image xlink:href="https://example.com/c.png"/></draw:frame></text:p>""",
            files: new Dictionary<string, byte[]> { ["Pictures/a.png"] = Png(), ["Pictures/b.svg"] = "<svg/>"u8.ToArray() },
            manifestExtra: """<manifest:file-entry manifest:full-path="Pictures/b.svg" manifest:media-type="image/svg+xml"/>""");

        Assert.Single(Blocks(bytes)[0].Content!);
        Assert.Equal(
            ["sidecar.notice.imagePosition", "sidecar.notice.imageType=image/svg+xml", "sidecar.notice.imageWithoutContent"],
            Noted(bytes));
    }

    [Fact]
    public void ThePageComesFromTheFirstBlocksMasterPage()
    {
        var bytes = Build(
            """<text:p text:style-name="P1">x</text:p>""",
            """<style:style style:name="P1" style:family="paragraph" style:master-page-name="Carta"/>""",
            pageLayouts: DefaultPage +
                         """<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="11in" fo:page-height="8.5in" fo:margin-top="10mm" fo:margin-bottom="15mm" fo:margin-left="30mm" fo:margin-right="20mm"/>""" +
                         """<style:header-style><style:header-footer-properties fo:min-height="5mm" fo:margin-bottom="5mm"/></style:header-style><style:footer-style/></style:page-layout>""" +
                         CenteredBandStyle,
            masters: DefaultMaster +
                     """<style:master-page style:name="Carta" style:page-layout-name="pm2"><style:header><text:p text:style-name="MP1"><text:tab/>Relatório — página <text:page-number text:select-page="current">1</text:page-number> de <text:page-count>3</text:page-count></text:p></style:header></style:master-page>""");

        var page = Open(bytes).Model.Page;
        Assert.Equal(("Letter", "landscape"), (page.Size, page.Orientation));
        Assert.Equal(new MarginsDto(20, 20, 15, 30), page.Margins);
        Assert.Equal(10, page.HeaderDistanceMm);
        Assert.Equal("Relatório — página {n} de {total}", page.HeaderText);
        Assert.Equal(string.Empty, page.FooterText);
        Assert.Empty(Noted(bytes));
    }

    private const string CenteredBandStyle =
        """<style:style style:name="MP1" style:family="paragraph" style:parent-style-name="Header"><style:paragraph-properties fo:text-align="center"/></style:style>""";

    [Theory]
    [InlineData("""<text:p>Relatório</text:p>""")]
    [InlineData("""<text:p text:style-name="MP2">Relatório</text:p>""")]
    [InlineData("""<text:p text:style-name="MP1">Relatório <text:span text:style-name="MT1">anual</text:span></text:p>""")]
    public void ABandThatIsNotPlainCenteredTextIsNoted(string paragraph)
    {
        var bytes = Build(
            "<text:p>x</text:p>",
            pageLayouts: DefaultPage + CenteredBandStyle +
                         """<style:style style:name="MP2" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-size="8pt"/></style:style>""" +
                         """<style:style style:name="MT1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>""",
            masters: $"""<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:footer>{paragraph}</style:footer></style:master-page>""");

        Assert.StartsWith("Relatório", Open(bytes).Model.Page.FooterText, StringComparison.Ordinal);
        Assert.Equal(["sidecar.notice.bandFormatting"], Noted(bytes));
    }

    [Fact]
    public void WhatThePageCannotCarryIsNoted()
    {
        var bytes = Build(
            """<text:p>x</text:p><text:p text:style-name="P1">y</text:p>""",
            """<style:style style:name="P1" style:family="paragraph" style:master-page-name="Outra"/>""",
            pageLayouts: """<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="148mm" fo:page-height="210mm"/></style:page-layout>""",
            masters: """<style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Data: <text:date>1/1/26</text:date></text:p><table:table/></style:header><style:footer-first><text:p>Capa</text:p></style:footer-first></style:master-page>""" +
                     """<style:master-page style:name="Outra" style:page-layout-name="pm1"/>""");

        var page = Open(bytes).Model.Page;
        Assert.Equal("A4", page.Size);
        Assert.Equal("Data: 1/1/26", page.HeaderText);
        Assert.Equal(
            ["sidecar.notice.bandFormatting", "sidecar.notice.headerFields", "sidecar.notice.headerVariants",
             "sidecar.notice.paperSizeChanged", "sidecar.notice.sections"],
            Noted(bytes));
    }

    [Fact]
    public void MetaBecomesTheProperties()
    {
        var properties = Open(Build(
            "<text:p/>",
            meta: """<meta:generator>LibreOffice</meta:generator><dc:title>Relatório</dc:title><dc:subject>Contas</dc:subject><dc:description>Resumo</dc:description>""" +
                  """<meta:keyword>contas</meta:keyword><meta:keyword>balanço</meta:keyword><meta:initial-creator>Ana</meta:initial-creator><dc:creator>Bia</dc:creator>""" +
                  """<meta:creation-date>2026-01-02T03:04:05</meta:creation-date><dc:date>2026-02-03T04:05:06</dc:date><meta:editing-cycles>4</meta:editing-cycles>""" +
                  """<meta:editing-duration>PT1H2M30S</meta:editing-duration><dc:language>pt-BR</dc:language><meta:user-defined meta:name="Company">ACME</meta:user-defined>""")).Model.Properties;

        Assert.Equal(
            new PropertiesDto(
                Title: "Relatório", Subject: "Contas", Creator: "Ana", Keywords: "contas; balanço", Description: "Resumo",
                LastModifiedBy: "Bia", Revision: "4", Created: "2026-01-02T03:04:05", Modified: "2026-02-03T04:05:06",
                Company: "ACME", Language: "pt-BR", TotalTime: 62),
            properties);
    }

    [Fact]
    public void WhatTheEditorDoesNotRepresentIsNotedAndAFieldWithoutANodeStaysAsText()
    {
        var bytes = Build(
            """<text:tracked-changes><text:changed-region text:id="c1"/></text:tracked-changes>""" +
            """<text:p>a<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>n</text:p></text:note-body></text:note>""" +
            """<text:note text:note-class="endnote"><text:note-body><text:p>e</text:p></text:note-body></text:note>""" +
            """<office:annotation><dc:creator>Bia</dc:creator><text:p>c</text:p></office:annotation>b<office:annotation-end/>""" +
            """<text:bookmark text:name="m"/><text:change text:change-id="c1"/><text:variable-get text:name="v">7</text:variable-get>""" +
            """<draw:frame><draw:object xlink:href="./Object 1"/></draw:frame><draw:frame><draw:text-box><text:p>t</text:p></draw:text-box></draw:frame>""" +
            """<draw:custom-shape/><draw:control/><text:ruby><text:ruby-base>r</text:ruby-base><text:ruby-text>x</text:ruby-text></text:ruby></text:p>""" +
            """<text:section text:name="S"><text:p>dentro</text:p></text:section>""" +
            """<text:table-of-content text:name="Sumário"><text:index-body><text:p>Capítulo</text:p></text:index-body></text:table-of-content>""" +
            """<office:forms><form:form/></office:forms>""",
            manifestExtra: """<manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.formula"/>""");

        var blocks = Blocks(bytes);
        Assert.Equal("ab7r", string.Concat(blocks[0].Content!.Select(node => node.Text)));
        Assert.Equal(["ab7r", "dentro", "Capítulo"], blocks.Select(block => string.Concat(block.Content!.Select(node => node.Text))));
        Assert.Equal(
            [
                "sidecar.notice.bookmarks", "sidecar.notice.comments", "sidecar.notice.drawings",
                "sidecar.notice.element=text:ruby", "sidecar.notice.endnotes", "sidecar.notice.equations",
                "sidecar.notice.footnotes", "sidecar.notice.forms", "sidecar.notice.indexes",
                "sidecar.notice.plainFields", "sidecar.notice.revisions", "sidecar.notice.sections",
                "sidecar.notice.textBoxes",
            ],
            Noted(bytes));
        Assert.Empty(Open(bytes).Inventory.Structural);
    }

    [Fact]
    public void AnEncryptedPackageIsRefused() =>
        AssertRefused(
            "sidecar.error.odtEncrypted",
            Build("<text:p/>", manifestExtra:
                """<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1"/></manifest:file-entry>"""));

    [Fact]
    public void APackageWithoutContentIsRefused() =>
        AssertRefused(
            "sidecar.error.odtWithoutContent",
            Zip([("mimetype", Encoding.ASCII.GetBytes(MimeType)), ("styles.xml", "<x/>"u8.ToArray())]));

    [Fact]
    public void ASpreadsheetIsNotAText() =>
        AssertRefused(
            "sidecar.error.odtNotText",
            Build("<text:p/>", mimeType: "application/vnd.oasis.opendocument.spreadsheet"));

    [Fact]
    public void ATextSavedFromHtmlIsRead()
    {
        var bytes = Build("<text:p>web</text:p>", mimeType: "application/vnd.oasis.opendocument.text-web");

        Assert.Contains("web", Roundtrip.TextOf(Open(bytes).Model), StringComparison.Ordinal);
    }

    [Fact]
    public void BytesThatAreNotAZipAreRefused() =>
        AssertRefused("sidecar.error.odtUnreadable", "não sou um zip"u8.ToArray());

    [Fact]
    public void MalformedXmlIsRefused() =>
        AssertRefused(
            "sidecar.error.odtUnreadable",
            Zip([("mimetype", Encoding.ASCII.GetBytes(MimeType)), ("content.xml", "<office:document-content"u8.ToArray())]));

    [Fact]
    public void ADocumentTypeDeclarationIsRefused() =>
        AssertRefused(
            "sidecar.error.odtUnreadable",
            Zip([("content.xml", """<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>"""u8.ToArray())]));

    [Fact]
    public void AnEntryPastTheCeilingIsRefusedWhateverItsHeaderSays()
    {
        var huge = new byte[OdtReader.MaxEntryBytes + 1];
        AssertRefused("sidecar.error.odtTooLarge", Zip([("content.xml", huge)]));
    }

    private static void AssertRefused(string key, byte[] bytes)
    {
        var problem = Assert.Throws<OdtException>(() => OdtReader.Read(bytes));
        Assert.Equal(key, problem.Message);
    }

    [Fact]
    public void FieldsWithAWordEquivalentBecomeFieldNodesAndTheDocumentStaysEditable()
    {
        var bytes = Build(
            """<text:p>p<text:page-number text:select-page="current">2</text:page-number>/<text:page-count>5</text:page-count> """ +
            """<text:date style:data-style-name="N30">06/11/26</text:date><text:time>10:00</text:time><text:title>Ata</text:title>""" +
            """<text:initial-creator>Ana</text:initial-creator><text:page-number text:select-page="next">3</text:page-number></text:p>""");

        var content = Blocks(bytes)[0].Content!;
        Assert.Equal(
            [" PAGE =2", " NUMPAGES =5", " DATE =06/11/26", " TIME =10:00", " TITLE =Ata", " AUTHOR =Ana"],
            content.Where(node => node.Type == "field").Select(node => $"{node.Attrs!["instr"]!.GetValue<string>()}={node.Attrs!["result"]!.GetValue<string>()}"));
        Assert.Equal("3", content[^1].Text);
        Assert.Equal(["sidecar.notice.plainFields"], Noted(bytes));
        Assert.Empty(Open(bytes).Inventory.Structural);
    }
}

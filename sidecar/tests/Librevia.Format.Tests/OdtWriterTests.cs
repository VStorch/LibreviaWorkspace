using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Librevia.Format.Docx;
using Librevia.Format.Odt;
using Librevia.Format.Protocol;
using static Librevia.Format.Tests.OdtFixtures;

namespace Librevia.Format.Tests;

/// <summary>
/// <c>odt.save</c>: what nobody touched goes back byte for byte, what was edited is written again,
/// and what an edited block held beyond the model is declared lost.
/// </summary>
public class OdtWriterTests
{
    /// <summary>The model as main sends it: through JSON, as the editor gives it back.</summary>
    private static DocumentModelDto ModelOf(byte[] odt) =>
        JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(OdtReader.Read(odt).Model, JsonOptions.Default), JsonOptions.Default)!;

    private static (byte[] Bytes, SaveResult Result) Save(byte[] odt, DocumentModelDto model) =>
        OdtWriter.Write(odt, JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(model, JsonOptions.Default), JsonOptions.Default)!);

    private static Dictionary<string, byte[]> Entries(byte[] zip)
    {
        using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return copy.ToArray();
        });
    }

    private static string Part(byte[] zip, string name) => Encoding.UTF8.GetString(Entries(zip)[name]);

    private static List<Node> Blocks(byte[] odt) => OdtReader.Read(odt).Model.Doc.Content!;

    private static string TextOf(Node node) =>
        node.Text ?? string.Concat((node.Content ?? []).Select(TextOf));

    private static List<string> Texts(byte[] odt) => [.. Blocks(odt).Select(TextOf)];

    private static Node Paragraph(string text) => Node.Of("paragraph", new Node { Type = "text", Text = text });

    private static string WithoutOid(Node node)
    {
        var copy = JsonSerializer.Deserialize<Node>(JsonSerializer.Serialize(node, DocxJson.Options), DocxJson.Options)!;
        copy.Attrs?.Remove("oid");
        return copy.Fingerprint();
    }

    private static List<string> Lost(SaveResult result) =>
        [.. result.Inventory.Lost.Select(notice => notice.Value is null ? notice.Key : $"{notice.Key}={notice.Value}")];

    [Fact]
    public void SavingWithoutEditsGivesBackEveryEntryWithTheSameContent()
    {
        var original = Rich();
        var (saved, result) = Save(original, ModelOf(original));

        var before = Entries(original);
        var after = Entries(saved);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        foreach (var (name, bytes) in before) Assert.True(bytes.AsSpan().SequenceEqual(after[name]), name);
        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(5, result.PreservedBlocks);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void WhatTheEditorAddsToEveryBlockIsNotAnEdit()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        foreach (var block in model.Doc.Content!)
        {
            block.With("indent", 0).With("floats", null).With("markRevision", null);
        }

        var (saved, result) = Save(original, model);

        Assert.Equal((3, 0), (result.PreservedBlocks, result.RewrittenBlocks));
        Assert.Equal(Entries(original)["content.xml"], Entries(saved)["content.xml"]);
    }

    [Fact]
    public void AnIndentLevelBecomesMillimetres()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        model.Doc.Content!.Add(Paragraph("recuado").With("indent", 2));

        var (saved, _) = Save(original, model);

        Assert.Equal("25.4", Blocks(saved)[^1].Attrs!["indentMm"]!.ToJsonString());
    }

    [Fact]
    public void MimetypeIsTheFirstEntryStoredWithoutExtraField()
    {
        var (saved, _) = Save(Rich(), ModelOf(Rich()));

        const int methodOffset = 8;
        const int nameLengthOffset = 26;
        const int extraLengthOffset = 28;
        const int headerLength = 30;
        Assert.Equal(0x04034b50, BitConverter.ToInt32(saved, 0));
        Assert.Equal(0, BitConverter.ToInt16(saved, methodOffset));
        Assert.Equal(0, BitConverter.ToInt16(saved, extraLengthOffset));
        var nameLength = BitConverter.ToInt16(saved, nameLengthOffset);
        Assert.Equal("mimetype", Encoding.ASCII.GetString(saved, headerLength, nameLength));
        Assert.Equal(MimeType, Encoding.ASCII.GetString(saved, headerLength + nameLength, MimeType.Length));
    }

    [Fact]
    public void EditingAParagraphChangesOnlyItsOwnTextInContent()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        model.Doc.Content![2].Content![0].Text = "Corpo novo";

        var (saved, result) = Save(original, model);

        var before = Part(original, "content.xml");
        var after = Part(saved, "content.xml");
        const string edited = """<text:p text:style-name="Text_20_body">Corpo</text:p>""";
        var start = before.IndexOf(edited, StringComparison.Ordinal);
        Assert.StartsWith(before[..start], after, StringComparison.Ordinal);
        Assert.EndsWith(before[(start + edited.Length)..], after, StringComparison.Ordinal);
        Assert.Contains("""<text:p text:style-name="Text_20_body">Corpo novo</text:p>""", after, StringComparison.Ordinal);
        Assert.Equal((2, 1), (result.PreservedBlocks, result.RewrittenBlocks));
        Assert.Equal(Entries(original)["styles.xml"], Entries(saved)["styles.xml"]);
    }

    [Fact]
    public void BlocksAreInsertedDeletedAndReordered()
    {
        var original = Rich();
        var model = ModelOf(original);
        var content = model.Doc.Content!;
        var first = content[0];
        content.RemoveAt(0);
        content.Insert(3, first);
        content.RemoveAt(content.Count - 1);
        content.Insert(1, Paragraph("Novo  com espaços"));

        var (saved, result) = Save(original, model);

        Assert.Equal(["umdois", "Novo  com espaços", "ab", "Com notanota aqui", "Primeiro"], Texts(saved));
        var text = Part(saved, "content.xml");
        Assert.True(text.IndexOf("text:sequence-decls", StringComparison.Ordinal) < text.IndexOf("Novo", StringComparison.Ordinal));
        Assert.Contains("<text:soft-page-break/>", text, StringComparison.Ordinal);
        Assert.Equal(4, result.PreservedBlocks);
        Assert.Equal(1, result.RewrittenBlocks);
    }

    [Fact]
    public void EditingAListItemOrACellRewritesOnlyThatListOrTable()
    {
        var original = Rich();
        var model = ModelOf(original);
        var list = model.Doc.Content![1];
        list.Content![1].Content![0].Content![0].Text = "três";
        var table = model.Doc.Content[2];
        table.Content![0].Content![1].Content![0].Content![0].Text = "B";

        var (saved, result) = Save(original, model);

        var blocks = Blocks(saved);
        Assert.Equal("umtrês", TextOf(blocks[1]));
        Assert.Equal("aB", TextOf(blocks[2]));
        Assert.Equal(WithoutOid(list), WithoutOid(blocks[1]));
        Assert.Equal(WithoutOid(table), WithoutOid(blocks[2]));
        Assert.Equal((3, 2), (result.PreservedBlocks, result.RewrittenBlocks));
        Assert.Contains("""<table:table table:name="Tabela1" """, Part(saved, "content.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void ANewTableKeepsItsMergedCells()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        Node Cell(string text, int colspan = 1, int rowspan = 1) =>
            Node.Of("tableCell", Paragraph(text)).With("colspan", colspan).With("rowspan", rowspan);
        var table = Node.Of(
            "table",
            Node.Of("tableRow", Cell("a", colspan: 2), Cell("c", rowspan: 2)),
            Node.Of("tableRow", Cell("d"), Cell("e")));
        model.Doc.Content!.Add(table);

        var (saved, _) = Save(original, model);

        var read = Blocks(saved)[^1];
        Assert.Equal("table", read.Type);
        var cells = read.Content!.Select(row => row.Content!.Select(cell =>
            $"{TextOf(cell)}:{cell.Attrs!["colspan"]}x{cell.Attrs["rowspan"]}")).ToList();
        Assert.Equal(["a:2x1", "c:1x2"], cells[0]);
        Assert.Equal(["d:1x1", "e:1x1"], cells[1]);
        Assert.Equal(2, Part(saved, "content.xml").Split("<table:covered-table-cell").Length - 1);
    }

    [Fact]
    public void ANewImageGoesToPicturesAndTheManifest()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        var src = $"data:image/png;base64,{Convert.ToBase64String(Png())}";
        var image = Node.Of("image").With("src", src).With("alt", "logo").With("width", 20).With("height", 30);
        model.Doc.Content!.Add(Node.Of("paragraph", image));

        var (saved, _) = Save(original, model);

        var entries = Entries(saved);
        Assert.Equal(Png(), entries["Pictures/Librevia1.png"]);
        Assert.Contains("""manifest:full-path="Pictures/Librevia1.png" manifest:media-type="image/png""", Part(saved, "META-INF/manifest.xml"), StringComparison.Ordinal);
        var read = Blocks(saved)[^1].Content![0];
        Assert.Equal(WithoutOid(image), WithoutOid(read));
    }

    [Fact]
    public void WithoutANewImageTheManifestStaysTheSame()
    {
        var original = Rich();
        var model = ModelOf(original);
        model.Doc.Content!.Add(Paragraph("fim"));

        var (saved, _) = Save(original, model);

        Assert.Equal(Entries(original)["META-INF/manifest.xml"], Entries(saved)["META-INF/manifest.xml"]);
    }

    [Fact]
    public void NewBlocksReadBackAsTheSameNodes()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        var bold = new Node { Type = "text", Text = "forte", Marks = [Mark.Of("bold"), Mark.Of("link", "href", "https://exemplo.org")] };
        var colored = new Node
        {
            Type = "text",
            Text = " e cor",
            Marks = [new Mark { Type = "textStyle", Attrs = new() { ["color"] = "#ff0000", ["fontSize"] = "14pt" } }],
        };
        var paragraph = Node.Of("paragraph", new Node { Type = "text", Text = " Antes\t" }, bold, colored, Node.Of("hardBreak"))
            .With("textAlign", "right").With("indentMm", 10).With("spaceAfter", 6);
        var heading = Node.Of("heading", new Node { Type = "text", Text = "Seção" }).With("level", 1).With("styleId", "Heading1");
        var nested = Node.Of("bulletList", Node.Of("listItem", Paragraph("dentro")));
        var list = Node.Of("orderedList", Node.Of("listItem", Paragraph("um"), nested), Node.Of("listItem", Paragraph("dois")));
        model.Doc.Content!.AddRange([paragraph, heading, list]);

        var (saved, _) = Save(original, model);

        var blocks = Blocks(saved);
        Assert.Equal(WithoutOid(paragraph), WithoutOid(blocks[3]));
        Assert.Equal(WithoutOid(heading), WithoutOid(blocks[4]));
        Assert.Equal(["umdentrodois"], [TextOf(blocks[5])]);
        Assert.Equal("orderedList", blocks[5].Type);
        Assert.Equal("bulletList", blocks[5].Content![0].Content![1].Type);
    }

    [Fact]
    public void ANamedStyleChangesOnlyInStyles()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        var body = model.Styles!.Styles["BodyText"];
        model.Styles.Styles["BodyText"] = body with { Paragraph = body.Paragraph! with { SpaceAfter = 12 } };

        var (saved, _) = Save(original, model);

        Assert.Equal(Entries(original)["content.xml"], Entries(saved)["content.xml"]);
        var styles = Part(saved, "styles.xml");
        Assert.Contains("""fo:margin-bottom="12pt""", styles, StringComparison.Ordinal);
        Assert.Contains("""fo:line-height="115%""", styles, StringComparison.Ordinal);
        Assert.Equal(12, OdtReader.Read(saved).Model.Styles!.Styles["BodyText"].Paragraph!.SpaceAfter);
        var before = Part(original, "styles.xml");
        Assert.Equal(before.Length - before.IndexOf("<office:automatic-styles>", StringComparison.Ordinal),
            styles.Length - styles.IndexOf("<office:automatic-styles>", StringComparison.Ordinal));
    }

    [Fact]
    public void ChangedMarginsRewriteOnlyThePageLayout()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        model = model with { Page = model.Page with { Margins = model.Page.Margins with { Left = 30 } } };

        var (saved, _) = Save(original, model);

        Assert.Equal(Entries(original)["content.xml"], Entries(saved)["content.xml"]);
        Assert.Contains("""fo:margin-left="30mm""", Part(saved, "styles.xml"), StringComparison.Ordinal);
        Assert.Equal(30, OdtReader.Read(saved).Model.Page.Margins.Left);
        Assert.Equal(20, OdtReader.Read(saved).Model.Page.Margins.Top);
    }

    [Fact]
    public void ANewHeaderTextBecomesTheMasterPageHeader()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        model = model with { Page = model.Page with { HeaderText = "Página {n} de {total}" } };

        var (saved, _) = Save(original, model);

        var page = OdtReader.Read(saved).Model.Page;
        Assert.Equal("Página {n} de {total}", page.HeaderText);
        Assert.Equal(20, page.Margins.Top);
    }

    [Fact]
    public void ChangedPropertiesRewriteOnlyMeta()
    {
        var original = Rich();
        var model = ModelOf(original);
        model = model with { Properties = model.Properties! with { Title = "Novo título", Subject = "Assunto" } };

        var (saved, _) = Save(original, model);

        var properties = OdtReader.Read(saved).Model.Properties!;
        Assert.Equal(("Novo título", "Assunto", "Ana"), (properties.Title, properties.Subject, properties.Creator));
        Assert.Equal(Entries(original)["content.xml"], Entries(saved)["content.xml"]);
        Assert.Equal(Entries(original)["styles.xml"], Entries(saved)["styles.xml"]);
    }

    [Fact]
    public void DeclarationsBeforeTheFirstBlockStayFirstWhenItIsDeleted()
    {
        var original = Rich();
        var model = ModelOf(original);
        model.Doc.Content!.RemoveAt(0);
        model.Doc.Content.Insert(0, Paragraph("Antes de tudo"));

        var (saved, _) = Save(original, model);

        var text = Part(saved, "content.xml");
        var body = text[text.IndexOf("<office:text>", StringComparison.Ordinal)..];
        Assert.StartsWith("<office:text><text:sequence-decls>", body, StringComparison.Ordinal);
        Assert.EndsWith("<text:soft-page-break/></office:text></office:body></office:document-content>", text, StringComparison.Ordinal);
    }

    [Fact]
    public void EditingAParagraphWithAFootnoteRewritesTheNoteWithItsId()
    {
        var original = Rich();
        Assert.Empty(OdtReader.Read(original).Inventory.Invisible);

        var model = ModelOf(original);
        model.Doc.Content![3].Content![0].Text = "Ainda com nota";
        var (saved, result) = Save(original, model);

        Assert.Empty(Lost(result));
        var text = Part(saved, "content.xml");
        Assert.Matches(
            """<text:p[^>]*>Ainda com nota<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p[^>]*>nota</text:p></text:note-body></text:note> aqui</text:p>""",
            text);
        var note = Blocks(saved)[3].Content![1];
        Assert.Equal(("noteRef", "nota"), (note.Type, TextOf(note)));
    }

    [Fact]
    public void OwnPrefixesAreUsedForWhatIsWritten()
    {
        var original = OwnPrefixes();
        var model = ModelOf(original);
        model.Doc.Content![1] = Paragraph("Três").With("textAlign", "center");

        var (saved, _) = Save(original, model);

        var text = Part(saved, "content.xml");
        Assert.Contains("<t:p>Um</t:p><t:p t:style-name=\"LvP1\">Três</t:p>", text, StringComparison.Ordinal);
        Assert.Contains("<o:automatic-styles><s:style s:name=\"LvP1\" s:family=\"paragraph\" s:parent-style-name=\"Standard\"><s:paragraph-properties f:text-align=\"center\"/></s:style></o:automatic-styles><o:body>", text, StringComparison.Ordinal);
        Assert.Equal(["Um", "Três"], Texts(saved));
        Assert.Equal("\"center\"", Blocks(saved)[1].Attrs!["textAlign"]!.ToJsonString());
    }

    [Fact]
    public void SavingTheSavedFileAgainChangesNothing()
    {
        var original = Rich();
        var model = ModelOf(original);
        model.Doc.Content![0].Content![0].Text = "Primeiro editado";
        model.Doc.Content.Add(Node.Of("table", Node.Of("tableRow", Node.Of("tableCell", Paragraph("x")).With("colspan", 1).With("rowspan", 1))));
        var (once, _) = Save(original, model);

        var (twice, result) = Save(once, ModelOf(once));

        var first = Entries(once);
        var second = Entries(twice);
        Assert.Equal(first.Keys.Order(), second.Keys.Order());
        foreach (var (name, bytes) in first) Assert.True(bytes.AsSpan().SequenceEqual(second[name]), name);
        Assert.Equal(0, result.RewrittenBlocks);
    }

    [Fact]
    public void APageBreakBeforeAWrittenParagraphGoesIntoItsStyle()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        model.Doc.Content!.AddRange([Node.Of("pageBreak"), Paragraph("Nova página")]);

        var (saved, result) = Save(original, model);

        var blocks = Blocks(saved);
        Assert.Equal(["pageBreak", "paragraph"], blocks[^2..].Select(block => block.Type));
        Assert.Equal(2, result.RewrittenBlocks);
    }

    [Fact]
    public void ABlockPastedTwiceIsKeptOnceAndWrittenTheSecondTime()
    {
        var original = LibreOffice();
        var model = ModelOf(original);
        var copy = JsonNode.Parse(JsonSerializer.Serialize(model.Doc.Content![2], JsonOptions.Default))!.Deserialize<Node>(JsonOptions.Default)!;
        model.Doc.Content.Add(copy);

        var (saved, result) = Save(original, model);

        Assert.Equal(["Título", "Olá mundo e  mais\tfimlinha forte", "Corpo", "Corpo"], Texts(saved));
        Assert.Equal((3, 1), (result.PreservedBlocks, result.RewrittenBlocks));
    }

    [Fact]
    public void EditedCellsKeepNoPartialAndFullBorders()
    {
        const string cells = """<table:table-cell table:style-name="{0}" office:value-type="string"><text:p>a</text:p></table:table-cell><table:table-cell table:style-name="{0}" office:value-type="string"><text:p>b</text:p></table:table-cell>""";
        string Table(string name, string style) =>
            $"""<table:table table:name="{name}"><table:table-column table:number-columns-repeated="2"/><table:table-row>{string.Format(cells, style)}</table:table-row></table:table>""";
        var original = Build(
            Table("Sem", "none") + Table("Parcial", "top") + Table("Completa", "full"),
            """<style:style style:name="top" style:family="table-cell"><style:table-cell-properties fo:border-top="0.5pt solid #000000" fo:border-bottom="none"/></style:style>""" +
            """<style:style style:name="full" style:family="table-cell"><style:table-cell-properties fo:border="1pt solid #ff0000"/></style:style>""");
        var model = ModelOf(original);
        var before = model.Doc.Content!.Select(WithoutOid).ToList();
        foreach (var table in model.Doc.Content!) table.Content![0].Content![0].Content![0].Content![0].Text = "editado";

        var (saved, result) = Save(original, model);

        var after = Blocks(saved);
        Assert.Equal((0, 3), (result.PreservedBlocks, result.RewrittenBlocks));
        Assert.All(after, table => Assert.Equal("editadob", TextOf(table)));
        foreach (var (table, index) in after.Select((table, index) => (table, index)))
        {
            Assert.Equal(before[index].Replace("\"a\"", "\"editado\""), WithoutOid(table));
        }

        var text = Part(saved, "content.xml");
        Assert.Contains("""fo:padding="0mm" fo:border-top="none" fo:border-right="none" fo:border-bottom="none" fo:border-left="none"/>""", text, StringComparison.Ordinal);
        Assert.Contains("""fo:border-top="0.5pt solid #000000" fo:border-right="none" fo:border-bottom="none" fo:border-left="none"/>""", text, StringComparison.Ordinal);
        Assert.Contains("""fo:border-top="1pt solid #ff0000" fo:border-right="1pt solid #ff0000" fo:border-bottom="1pt solid #ff0000" fo:border-left="1pt solid #ff0000"/>""", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEditedParagraphKeepsTheStandardStyle()
    {
        var original = Build("""<text:p text:style-name="Standard">Antes</text:p>""");
        var model = ModelOf(original);
        model.Doc.Content![0].Content![0].Text = "Depois";

        var (saved, _) = Save(original, model);

        Assert.Contains("""<text:p text:style-name="Standard">Depois</text:p>""", Part(saved, "content.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void AFieldInAnEditedParagraphGoesBackAsTheSameOdfField()
    {
        const string date = """<text:date style:data-style-name="N30" text:date-value="2026-11-06T09:56:12" text:fixed="true">06/11/26</text:date>""";
        var original = Build($"""<text:p text:style-name="Standard">Hoje {date}, página <text:page-number text:select-page="current">1</text:page-number></text:p>""");
        var model = ModelOf(original);
        model.Doc.Content![0].Content![0].Text = "Agora ";

        var (saved, result) = Save(original, model);

        var text = Part(saved, "content.xml");
        Assert.Contains($"""<text:p text:style-name="Standard">Agora {date}, página <text:page-number text:select-page="current">1</text:page-number></text:p>""", text, StringComparison.Ordinal);
        Assert.Empty(Lost(result));
        Assert.Equal([" DATE ", " PAGE "], Blocks(saved)[0].Content!.Where(node => node.Type == "field").Select(node => node.Attrs!["instr"]!.GetValue<string>()));
    }

    [Fact]
    public void AFieldWithoutAnOdfEquivalentKeepsItsTextAndIsDeclared()
    {
        var original = Build("<text:p>a</text:p>");
        var model = ModelOf(original);
        model.Doc.Content![0].Content!.Add(Node.Of("field").With("instr", "SEQ Figura").With("result", "3"));
        model.Doc.Content!.Add(Paragraph("x"));
        model.Doc.Content![1].Content!.Add(Node.Of("field").With("instr", "NUMPAGES").With("result", "9"));

        var (saved, result) = Save(original, model);

        var text = Part(saved, "content.xml");
        Assert.Contains(">a3</text:p>", text, StringComparison.Ordinal);
        Assert.Contains("<text:page-count>9</text:page-count>", text, StringComparison.Ordinal);
        Assert.Equal(["sidecar.notice.fields"], Lost(result));
    }

    [Fact]
    public void AnEmbeddedObjectWhoseLastFrameIsGoneLeavesThePackage()
    {
        const string frame = """<draw:frame draw:name="Eq{0}"><draw:object xlink:href="./Object {0}"/><draw:image xlink:href="./ObjectReplacements/Object {0}"/></draw:frame>""";
        var math = Utf8("""<math xmlns="http://www.w3.org/1998/Math/MathML"/>""");
        var original = Build(
            $"<text:p>um{string.Format(frame, 1)}</text:p><text:p>dois{string.Format(frame, 2)}</text:p>",
            files: new Dictionary<string, byte[]>
            {
                ["Object 1/content.xml"] = math,
                ["ObjectReplacements/Object 1"] = Png(),
                ["Object 2/content.xml"] = math,
                ["ObjectReplacements/Object 2"] = Png(),
            });
        var model = ModelOf(original);
        model.Doc.Content![0] = Paragraph("um");

        var (saved, _) = Save(original, model);

        var entries = Entries(saved).Keys;
        Assert.DoesNotContain("Object 1/content.xml", entries);
        Assert.DoesNotContain("ObjectReplacements/Object 1", entries);
        Assert.Contains("Object 2/content.xml", entries);
        Assert.Contains("ObjectReplacements/Object 2", entries);
        var manifest = Part(saved, "META-INF/manifest.xml");
        Assert.DoesNotContain("\"Object 1/", manifest, StringComparison.Ordinal);
        Assert.DoesNotContain("ObjectReplacements/Object 1\"", manifest, StringComparison.Ordinal);
        Assert.Contains("\"Object 2/\"", manifest, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(" leading")]
    [InlineData("tab\t after")]
    [InlineData("three   spaces")]
    [InlineData("trailing  ")]
    public void WhiteSpaceThatOdfWouldCollapseSurvivesARewrite(string text)
    {
        var original = Build("""<text:p text:style-name="Standard">a</text:p>""");
        var model = ModelOf(original);
        model.Doc.Content![0].Content![0].Text = text;

        var (saved, _) = Save(original, model);

        Assert.Equal([text], Texts(saved));
    }

    [Fact]
    public void AFieldResultReadsItsWhiteSpaceAsShown()
    {
        var original = Build("""<text:p text:style-name="Standard"><text:page-number text:select-page="current">1 &#10;   <text:s/>de<text:tab/>2</text:page-number></text:p>""");

        var field = Blocks(original)[0].Content!.Single(node => node.Type == "field");

        Assert.Equal("1  de\t2", field.Attrs!["result"]!.GetValue<string>());
    }
}

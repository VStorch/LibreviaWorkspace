using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Librevia.Format.Docx;
using Librevia.Format.Odt;
using Librevia.Format.Protocol;
using static Librevia.Format.Tests.OdtFixtures;

namespace Librevia.Format.Tests;

/// <summary>
/// Page setup by section: a block whose style opens another master page starts a section on a new
/// page, a <c>text:section</c> with columns a continuous one; saving maps them back to the file's own
/// master pages.
/// </summary>
public class OdtSectionTests
{
    private const string Landscape =
        """<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="29.7cm" fo:page-height="21.001cm" style:print-orientation="landscape" fo:margin-top="1cm" fo:margin-bottom="1cm" fo:margin-left="1cm" fo:margin-right="1cm"/></style:page-layout>""";

    private const string Masters =
        DefaultMaster + """<style:master-page style:name="Paisagem" style:page-layout-name="pm2"/>""";

    private const string Automatic =
        """<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard" style:master-page-name="Paisagem"><style:paragraph-properties style:page-number="5"/></style:style>""" +
        """<style:style style:name="P2" style:family="paragraph" style:parent-style-name="Standard" style:master-page-name="Standard"><style:paragraph-properties style:page-number="auto"/></style:style>""" +
        """<style:style style:name="Sect1" style:family="section"><style:section-properties><style:columns fo:column-count="2" fo:column-gap="5mm"><style:column-sep style:width="0.2mm"/></style:columns></style:section-properties></style:style>""";

    /// <summary>Portrait, then landscape from page 5, then two columns, then one again.</summary>
    private static byte[] Sectioned() => Build(
        """<text:p text:style-name="Standard">Retrato</text:p><text:p text:style-name="Standard">Fim do retrato</text:p>""" +
        """<text:p text:style-name="P1">Paisagem</text:p><text:p text:style-name="Standard">Mais paisagem</text:p>""" +
        """<text:section text:style-name="Sect1" text:name="Colunas" text:protected="false"><text:p>Coluna A</text:p><text:p>Coluna B</text:p></text:section>""" +
        """<text:p text:style-name="Standard">Depois</text:p>""",
        automatic: Automatic, pageLayouts: DefaultPage + Landscape, masters: Masters);

    /// <summary>Portrait, landscape, portrait again on the same master page.</summary>
    private static byte[] BackToPortrait() => Build(
        """<text:p>Um</text:p><text:p text:style-name="P1">Dois</text:p><text:p text:style-name="P2">Três</text:p>""",
        automatic: Automatic, pageLayouts: DefaultPage + Landscape, masters: Masters);

    private static DocumentModelDto ModelOf(byte[] odt) => RoundTrip(OdtReader.Read(odt).Model);

    private static DocumentModelDto RoundTrip(DocumentModelDto model) =>
        JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(model, JsonOptions.Default), JsonOptions.Default)!;

    private static (byte[] Bytes, SaveResult Result) Save(byte[] odt, DocumentModelDto model) => OdtWriter.Write(odt, RoundTrip(model));

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

    private static string TextOf(Node node) => node.Text ?? string.Concat((node.Content ?? []).Select(TextOf));

    private static string? BreakOf(Node node) =>
        (node.Attrs?.TryGetValue("sectionBreak", out var own) == true ? own?.GetValue<string>() : null) ??
        (node.Content ?? []).Select(BreakOf).FirstOrDefault(found => found is not null);

    private static List<string> Breaks(DocumentModelDto model) =>
        [.. model.Doc.Content!.Where(node => BreakOf(node) is not null).Select(node => $"{TextOf(node)}={BreakOf(node)}")];

    private static List<PageSetupDto> All(DocumentModelDto model) => [.. model.Sections ?? [], model.Page];

    private static int? StartOf(PageSetupDto page) => PageReader.TryStartOf(page, out var start) ? start : null;

    private static string Shape(PageSetupDto page) =>
        $"{page.Size} {page.Orientation} {page.Margins.Left} {page.Start} {StartOf(page)} {page.Columns?.Count}";

    private static List<string> Noted(byte[] odt) => [.. OdtReader.Read(odt).Inventory.Invisible.Select(notice => notice.Key)];

    [Fact]
    public void EachMasterPageTheBlocksOpenIsASection()
    {
        var model = ModelOf(Sectioned());

        Assert.Equal(
            ["A4 portrait 20 nextPage  ", "A4 landscape 10 nextPage 5 ", "A4 landscape 10 continuous  2", "A4 landscape 10 continuous  "],
            All(model).Select(Shape));
        Assert.Equal(["s1", "s2", "s3"], model.Sections!.Select(section => section.Id));
        Assert.Equal(["Fim do retrato=s1", "Mais paisagem=s2", "Coluna B=s3"], Breaks(model));
        Assert.DoesNotContain("sidecar.notice.sections", Noted(Sectioned()));
    }

    [Fact]
    public void ATextSectionWithColumnsIsAContinuousSection()
    {
        var columns = ModelOf(Sectioned()).Sections![2].Columns!;

        Assert.Equal((2, 5.0, true), (columns.Count, columns.SpaceMm, columns.Separator));
    }

    [Fact]
    public void ThePageNumberRestartsWhereTheMasterPageSaysSo()
    {
        var model = ModelOf(BackToPortrait());

        Assert.Equal([null, 5, null], All(model).Select(StartOf));
        Assert.Equal(["A4 portrait 20 nextPage  ", "A4 landscape 10 nextPage 5 ", "A4 portrait 20 nextPage  "], All(model).Select(Shape));
    }

    [Fact]
    public void ANestedMasterPageOrSectionKeepsTheNotice()
    {
        var nested = Build(
            """<text:p>a</text:p><text:list><text:list-item><text:p>b</text:p></text:list-item><text:list-item><text:p text:style-name="P1">c</text:p></text:list-item></text:list>""" +
            """<text:section text:style-name="Sect1" text:name="S"><text:section text:name="T"><text:p>d</text:p></text:section></text:section>""",
            automatic: Automatic, pageLayouts: DefaultPage + Landscape, masters: Masters);

        Assert.Contains("sidecar.notice.sections", Noted(nested));
        Assert.Null(ModelOf(nested).Sections);
    }

    [Fact]
    public void SavingWithoutEditsGivesBackTheSameBytes()
    {
        foreach (var original in new[] { Sectioned(), BackToPortrait() })
        {
            var (saved, result) = Save(original, ModelOf(original));

            var before = Entries(original);
            foreach (var (name, bytes) in Entries(saved)) Assert.True(bytes.AsSpan().SequenceEqual(before[name]), name);
            Assert.Equal(0, result.RewrittenBlocks);
            Assert.Empty(result.Inventory.Lost);
        }
    }

    [Fact]
    public void EditingTheBlockThatOpensASectionKeepsTheSections()
    {
        var original = Sectioned();
        var model = ModelOf(original);
        model.Doc.Content![2].Content![0].Text = "Paisagem editada";

        var (saved, result) = Save(original, model);

        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);
        Assert.Matches("style:master-page-name=\"Paisagem\"><style:paragraph-properties style:page-number=\"5\"/>", Part(saved, "content.xml"));
        Assert.Equal(All(model).Select(Shape), All(ModelOf(saved)).Select(Shape));
        Assert.Equal(Breaks(model), Breaks(ModelOf(saved)));
    }

    [Fact]
    public void EditingInsideANamedColumnsSectionKeepsItsTag()
    {
        var original = Sectioned();
        var model = ModelOf(original);
        model.Doc.Content![4].Content![0].Text = "Coluna A editada";

        var (saved, result) = Save(original, model);
        var content = Part(saved, "content.xml");

        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("<text:section text:style-name=\"Sect1\" text:name=\"Colunas\" text:protected=\"false\">", content);
        Assert.DoesNotContain("LvSect", content);
        Assert.Equal(All(model).Select(Shape), All(ModelOf(saved)).Select(Shape));
    }

    [Fact]
    public void ChangingASectionsMarginsChangesItsMasterPageOnly()
    {
        var original = Sectioned();
        var model = ModelOf(original);
        model = model with { Sections = [model.Sections![0] with { Margins = new MarginsDto(20, 20, 20, 30) }, .. model.Sections.Skip(1)] };

        var (saved, result) = Save(original, model);

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.True(Entries(saved)["content.xml"].AsSpan().SequenceEqual(Entries(original)["content.xml"]));
        Assert.Equal([30, 10, 10, 10], All(ModelOf(saved)).Select(page => page.Margins.Left));
    }

    [Fact]
    public void ASharedMasterPageWithAnotherSetupIsCopied()
    {
        var original = BackToPortrait();
        var model = ModelOf(original);
        model = model with { Page = model.Page with { Margins = new MarginsDto(20, 20, 20, 35) } };

        var (saved, result) = Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Contains("style:master-page-name=\"LvPage1\"", Part(saved, "content.xml"));
        Assert.Contains("<style:master-page style:name=\"LvPage1\" style:page-layout-name=\"LvPm1\"", Part(saved, "styles.xml"));
        Assert.Equal([20, 10, 35], All(ModelOf(saved)).Select(page => page.Margins.Left));
    }

    [Fact]
    public void DeletingASectionBreakJoinsTheSectionToTheNext()
    {
        var original = Sectioned();
        var model = ModelOf(original);
        model.Doc.Content![1].Attrs!.Remove("sectionBreak");
        model = model with { Sections = [.. model.Sections!.Skip(1)] };

        var (saved, _) = Save(original, model);
        var read = ModelOf(saved);

        Assert.Equal(
            ["A4 landscape 10 nextPage 5 ", "A4 landscape 10 continuous  2", "A4 landscape 10 continuous  "],
            All(read).Select(Shape));
        Assert.Equal(["Mais paisagem=s1", "Coluna B=s2"], Breaks(read));
        Assert.Equal(
            ["Retrato", "Fim do retrato", "Paisagem", "Mais paisagem", "Coluna A", "Coluna B", "Depois"],
            read.Doc.Content!.Select(TextOf));
    }

    [Fact]
    public void ANewDocumentReadsBackTheSectionsItWasCreatedWith()
    {
        var none = PageReader.StartElement(null);
        var portrait = new PageSetupDto("A4", "portrait", new MarginsDto(20, 20, 20, 20), null, null, PageNumberFormat: "decimal", PageNumberStart: none, Start: "nextPage");
        var landscape = new PageSetupDto("Letter", "landscape", new MarginsDto(15, 15, 15, 15), null, null, PageNumberFormat: "decimal", PageNumberStart: none, Start: "nextPage");
        List<PageSetupDto> sections =
        [
            portrait with { Id = "a" },
            landscape with { Id = "b", PageNumberStart = PageReader.StartElement(3) },
            landscape with { Id = "c", Start = "continuous", Columns = new ColumnsDto(2, 10, true) },
        ];
        var page = landscape with { Start = "continuous" };
        var doc = Node.Of(
            "doc",
            Node.Of("paragraph", new Node { Type = "text", Text = "Um" }).With("sectionBreak", "a"),
            Node.Of("paragraph", new Node { Type = "text", Text = "Dois" }).With("sectionBreak", "b"),
            Node.Of("paragraph", new Node { Type = "text", Text = "Três" }).With("sectionBreak", "c"),
            Node.Of("paragraph", new Node { Type = "text", Text = "Quatro" }));
        var model = new DocumentModelDto(page, doc, Sections: sections);

        var template = OdtTemplate.Create(new OdtCreateDto(page, sections, null, null, "Librevia", null, []));
        var (created, result) = OdtWriter.Write(template, RoundTrip(model));
        var read = ModelOf(created);

        Assert.Empty(result.Inventory.Lost);
        Assert.Equal(All(model).Select(Shape), All(read).Select(Shape));
        Assert.Equal(["Um=s1", "Dois=s2", "Três=s3"], Breaks(read));
    }
}

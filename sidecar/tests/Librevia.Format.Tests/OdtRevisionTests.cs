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
/// ODF tracked changes as DOCX revisions: insertion ranges and deletion points become the
/// <c>insertion</c> and <c>deletion</c> marks and the paragraph mark's <c>markRevision</c>, and
/// saving writes them back with their regions.
/// </summary>
public class OdtRevisionTests
{
    private const string Regions =
        """<text:tracked-changes>""" +
        """<text:changed-region xml:id="ct1" text:id="ct1"><text:insertion><office:change-info><dc:creator>Ana</dc:creator><dc:date>2024-01-02T10:00:00</dc:date></office:change-info></text:insertion></text:changed-region>""" +
        """<text:changed-region xml:id="ct2" text:id="ct2"><text:deletion><office:change-info><dc:creator>Bruno</dc:creator><dc:date>2024-01-03T11:00:00</dc:date></office:change-info><text:p text:style-name="Standard">removido</text:p></text:deletion></text:changed-region>""" +
        """<text:changed-region xml:id="ct3" text:id="ct3"><text:deletion><office:change-info><dc:creator>Bruno</dc:creator><dc:date>2024-01-03T11:00:00</dc:date></office:change-info><text:p text:style-name="Standard"/><text:p text:style-name="Text_20_body"/></text:deletion></text:changed-region>""" +
        """<text:changed-region xml:id="ct4" text:id="ct4"><text:insertion><office:change-info><dc:creator>Ana</dc:creator><dc:date>2024-01-04T09:00:00</dc:date></office:change-info></text:insertion></text:changed-region>""" +
        """</text:tracked-changes>""";

    /// <summary>What LibreOffice writes for a DOCX with an insertion, a deletion and a deleted paragraph mark.</summary>
    private static byte[] Tracked() => Build(
        Regions +
        """<text:p text:style-name="Standard">Alfa <text:change-start text:change-id="ct1"/>inserido<text:change-end text:change-id="ct1"/> beta <text:change text:change-id="ct2"/><text:s/>gama.</text:p>""" +
        """<text:p text:style-name="Standard">Primeira<text:change text:change-id="ct3"/>segunda.</text:p>""" +
        """<text:p text:style-name="Standard">Antes<text:change-start text:change-id="ct4"/></text:p>""" +
        """<text:p text:style-name="Standard">novo<text:change-end text:change-id="ct4"/></text:p>""" +
        """<text:p text:style-name="Standard">Intocado.</text:p>""");

    private static DocumentModelDto ModelOf(byte[] odt) =>
        JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(OdtReader.Read(odt).Model, JsonOptions.Default), JsonOptions.Default)!;

    private static List<Node> Blocks(byte[] odt) => ModelOf(odt).Doc.Content!;

    private static byte[] Save(byte[] odt, DocumentModelDto model) => OdtWriter.Write(odt, model).Bytes;

    private static string Content(byte[] odt)
    {
        using var archive = new ZipArchive(new MemoryStream(odt), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry("content.xml")!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static string? Value(IDictionary<string, JsonNode?>? attrs, string name) =>
        attrs?.TryGetValue(name, out var value) == true ? value?.GetValue<string>() : null;

    private static string? StringAttr(Node node, string name) => Value(node.Attrs, name);

    private static string Revision(Node node) =>
        string.Join(",", (node.Marks ?? []).Where(mark => mark.Type is Revisions.Insertion or Revisions.Deletion)
            .Select(mark => $"{mark.Type}:{Value(mark.Attrs, "author")}:{Value(mark.Attrs, "rid")}"));

    private static List<string> Runs(Node block) =>
        [.. (block.Content ?? []).Select(node => $"{node.Text}|{Revision(node)}")];

    private static string? MarkRevision(Node block) =>
        block.Attrs?.GetValueOrDefault("markRevision") is JsonObject revision
            ? $"{revision["kind"]}:{revision["author"]}:{revision["rid"]}"
            : null;

    private static JsonObject BreakOf(Mark mark) => new()
    {
        ["kind"] = mark.Type == Revisions.Deletion ? "del" : "ins",
        ["author"] = Value(mark.Attrs, "author"),
        ["date"] = Value(mark.Attrs, "date"),
    };

    private static Node Text(string text, params Mark[] marks) =>
        new() { Type = "text", Text = text, Marks = marks.Length == 0 ? null : [.. marks] };

    private static Mark Revised(string type, string author, string date) =>
        new() { Type = type, Attrs = new() { ["author"] = author, ["date"] = date } };

    [Fact]
    public void InsertionsAndDeletionsBecomeRevisionMarksWithTheirAuthorDateAndRegion()
    {
        var blocks = Blocks(Tracked());

        Assert.Equal(
            ["Alfa |", "inserido|insertion:Ana:ct1", " beta |", "removido|deletion:Bruno:ct2", " gama.|"],
            Runs(blocks[0]));
        var inserted = blocks[0].Content![1].Marks!.Single();
        Assert.Equal("2024-01-02T10:00:00", Value(inserted.Attrs, "date"));
        Assert.True(ModelOf(Tracked()).TrackChanges);
    }

    [Fact]
    public void ADeletedParagraphMarkSplitsTheParagraphAndTheRestTakesTheStoredStyle()
    {
        var blocks = Blocks(Tracked());

        Assert.Equal(["Primeira|"], Runs(blocks[1]));
        Assert.Equal("del:Bruno:ct3", MarkRevision(blocks[1]));
        Assert.Equal(["segunda.|"], Runs(blocks[2]));
        Assert.Null(MarkRevision(blocks[2]));
        Assert.Equal("BodyText", StringAttr(blocks[2], "styleId"));
        Assert.Equal(blocks[1].Attrs!["oid"]!.GetValue<string>() + ".2", StringAttr(blocks[2], "oid"));
    }

    [Fact]
    public void AnInsertionAcrossParagraphsInsertsTheParagraphMarkItCovers()
    {
        var blocks = Blocks(Tracked());

        Assert.Equal("ins:Ana:ct4", MarkRevision(blocks[3]));
        Assert.Equal(["novo|insertion:Ana:ct4"], Runs(blocks[4]));
        Assert.Null(MarkRevision(blocks[4]));
    }

    [Fact]
    public void AFormatChangeIsNotedAndItsTextStays()
    {
        var bytes = Build(
            """<text:tracked-changes><text:changed-region text:id="f1"><text:format-change><office:change-info><dc:creator>Ana</dc:creator><dc:date>2024-01-01T00:00:00</dc:date></office:change-info></text:format-change></text:changed-region></text:tracked-changes>""" +
            """<text:p>a<text:change-start text:change-id="f1"/>b<text:change-end text:change-id="f1"/></text:p>""");

        var result = OdtReader.Read(bytes);

        Assert.Equal(["ab|"], Runs(result.Model.Doc.Content![0]));
        Assert.Contains(result.Inventory.Invisible, notice => notice.Key == "sidecar.notice.revisions");
    }

    [Fact]
    public void WithoutTrackedChangesOrWithTheFlagOffRecordingIsOff()
    {
        Assert.Null(ModelOf(Build("<text:p>a</text:p>")).TrackChanges);
        Assert.Null(ModelOf(Build("""<text:tracked-changes text:track-changes="false"/><text:p>a</text:p>""")).TrackChanges);
    }

    [Fact]
    public void SavingWithoutEditsGivesBackTheSameContent()
    {
        var odt = Tracked();

        Assert.Equal(Content(odt), Content(Save(odt, ModelOf(odt))));
    }

    [Fact]
    public void EditingNextToARevisionKeepsItAndTheOtherBlocksRegions()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        model.Doc.Content![0].Content![0].Text = "Alfa editado ";

        var saved = Save(odt, model);

        var blocks = Blocks(saved);
        Assert.Equal(
            ["Alfa editado |", "inserido|insertion:Ana:ct1", " beta |", "removido|deletion:Bruno:ct2", " gama.|"],
            Runs(blocks[0]));
        Assert.Equal("del:Bruno:ct3", MarkRevision(blocks[1]));
        Assert.Equal("ins:Ana:ct4", MarkRevision(blocks[3]));
        var content = Content(saved);
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(content, "text:id=\"ct1\""));
        Assert.Contains("""<text:p text:style-name="Standard">Primeira<text:change text:change-id="ct3"/>segunda.</text:p>""", content);
        Assert.StartsWith("<text:tracked-changes>", content[(content.IndexOf("<office:text>", StringComparison.Ordinal) + "<office:text>".Length)..]);
    }

    [Fact]
    public void AcceptingAndRejectingLeavesNoRegionNobodyPointsTo()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        var first = model.Doc.Content![0];
        first.Content!.RemoveAt(3);
        first.Content[1].Marks = null;

        var content = Content(Save(odt, model));

        Assert.DoesNotContain("ct1", content);
        Assert.DoesNotContain("ct2", content);
        Assert.DoesNotContain("removido", content);
        Assert.Contains("text:id=\"ct3\"", content);
        Assert.Contains("text:id=\"ct4\"", content);
    }

    [Fact]
    public void AcceptingEveryRevisionDropsTheRegionsAndKeepsRecordingOn()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        var plain = new List<Node>
        {
            Node.Of("paragraph", Text("Alfa inserido beta gama.")),
            Node.Of("paragraph", Text("Primeirasegunda.")),
            Node.Of("paragraph", Text("Antes")),
            Node.Of("paragraph", Text("novo")),
            model.Doc.Content![^1],
        };
        model = model with { Doc = Node.Of("doc", [.. plain]) };

        var content = Content(Save(odt, model));

        Assert.DoesNotContain("changed-region", content);
        Assert.DoesNotContain("text:change", content);
        Assert.Contains("<text:tracked-changes></text:tracked-changes>", content);
    }

    [Fact]
    public void ANewRevisionGetsARegionWhoseIdTheFileDoesNotUse()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        var mark = Revised(Revisions.Insertion, "Caio", "2026-11-06T12:00:00Z");
        model.Doc.Content![^1].Content = [Text("Intocado"), Text(" agora", mark), Text(".")];

        var saved = Save(odt, model);

        var last = Blocks(saved)[^1];
        Assert.Equal(["Intocado|", " agora|insertion:Caio:LvChange1", ".|"], Runs(last));
        Assert.Contains("text:id=\"ct1\"", Content(saved));
    }

    [Fact]
    public void ARewrittenDeletedParagraphMarkJoinsTheParagraphsAgain()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        model.Doc.Content![2].Content![0].Text = "segunda editada.";

        var blocks = Blocks(Save(odt, model));

        Assert.Equal(["Primeira|"], Runs(blocks[1]));
        Assert.Equal("del:Bruno:ct3", MarkRevision(blocks[1]));
        Assert.Equal(["segunda editada.|"], Runs(blocks[2]));
        Assert.Equal("BodyText", StringAttr(blocks[2], "styleId"));
    }

    [Fact]
    public void ARangeAcrossBlocksIsRewrittenWholeWhenOneSideIsEdited()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        model.Doc.Content![4].Content![0].Text = "novíssimo";

        var saved = Save(odt, model);

        var blocks = Blocks(saved);
        Assert.Equal("ins:Ana:ct4", MarkRevision(blocks[3]));
        Assert.Equal(["novíssimo|insertion:Ana:ct4"], Runs(blocks[4]));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(Content(saved), "change-start text:change-id=\"ct4\""));
    }

    /// <summary>A range from the first paragraph to the third: the middle one holds no mark of its own.</summary>
    private static byte[] Spanning() => Build(
        """<text:tracked-changes><text:changed-region xml:id="m1" text:id="m1"><text:insertion><office:change-info><dc:creator>Ana</dc:creator><dc:date>2024-01-01T00:00:00</dc:date></office:change-info></text:insertion></text:changed-region></text:tracked-changes>""" +
        """<text:p>Um<text:change-start text:change-id="m1"/> início</text:p><text:p>meio inteiro</text:p><text:p>fim<text:change-end text:change-id="m1"/> resto.</text:p>""");

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void EditingAnyParagraphOfARangeKeepsTheWholeRangeAndOnlyItsText(int edited)
    {
        var odt = Spanning();
        var model = ModelOf(odt);
        model.Doc.Content![edited].Content!.Add(Text(" X"));

        var saved = Save(odt, model);

        var blocks = Blocks(saved);
        var runs = blocks.Select(block => string.Join(" / ", Runs(block).Select(WithoutRegion))).ToList();
        var expected = new List<string>
        {
            "Um| /  início|insertion:Ana",
            "meio inteiro|insertion:Ana",
            "fim|insertion:Ana /  resto.|",
        };
        if (edited == 2) expected[2] = "fim|insertion:Ana /  resto. X|";
        else expected[edited] += " /  X|";
        Assert.Equal(expected, runs);
        Assert.StartsWith("ins:Ana:", MarkRevision(blocks[0]));
        Assert.StartsWith("ins:Ana:", MarkRevision(blocks[1]));
        var content = Content(saved);
        var regions = System.Text.RegularExpressions.Regex.Matches(content, "text:changed-region xml:id=\"([^\"]+)\"").Select(match => match.Groups[1].Value);
        Assert.All(regions, id => Assert.Contains($"change-start text:change-id=\"{id}\"", content));
    }

    private static string WithoutRegion(string run) => System.Text.RegularExpressions.Regex.Replace(run, ":(m1|LvChange\\d+)$", string.Empty);

    [Fact]
    public void InsertedTextBeforeAnInsertedParagraphMarkNobodyContinuesStaysInserted()
    {
        var odt = Tracked();
        var model = ModelOf(odt);
        var mark = Revised(Revisions.Insertion, "Caio", "2026-11-06T12:00:00Z");
        var first = model.Doc.Content![0];
        first.Content!.Add(Text(" novo", mark));
        first.With("markRevision", BreakOf(mark));

        var saved = Save(odt, model);

        var blocks = Blocks(saved);
        Assert.Equal(" novo|insertion:Caio:LvChange1", Runs(blocks[0])[^1]);
        Assert.Equal(["Primeira|"], Runs(blocks[1]));
    }

    [Fact]
    public void TurningRecordingOffWritesTheFlag()
    {
        var odt = Tracked();
        var model = ModelOf(odt) with { TrackChanges = false };

        var saved = Save(odt, model);

        Assert.Contains("""<text:tracked-changes text:track-changes="false">""", Content(saved));
        Assert.Null(ModelOf(saved).TrackChanges);
    }

    [Fact]
    public void ANewDocumentWritesItsRevisionsAndRecording()
    {
        var mark = Revised(Revisions.Deletion, "Ana", "2026-11-06T12:00:00Z");
        var inserted = Revised(Revisions.Insertion, "Bia", "2026-11-06T12:00:00Z");
        var first = Node.Of("paragraph", Text("um "), Text("velho", mark), Text(" novo", inserted));
        first.With("markRevision", BreakOf(inserted));
        var doc = Node.Of("doc", first, Node.Of("paragraph", Text("dois", inserted)));
        var blank = Build("");
        var model = ModelOf(blank) with { Doc = doc, TrackChanges = true };

        var saved = Save(blank, model);

        var blocks = Blocks(saved);
        Assert.Equal(["um |", "velho|deletion:Ana:LvChange1", " novo|insertion:Bia:LvChange2"], Runs(blocks[0]));
        Assert.Equal("ins:Bia:LvChange2", MarkRevision(blocks[0]));
        Assert.Equal(["dois|insertion:Bia:LvChange2"], Runs(blocks[1]));
        Assert.True(ModelOf(saved).TrackChanges);
    }
}

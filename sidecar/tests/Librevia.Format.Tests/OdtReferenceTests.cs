using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Librevia.Format.Docx;
using Librevia.Format.Odt;
using Librevia.Format.Protocol;
using static Librevia.Format.Tests.OdtFixtures;

namespace Librevia.Format.Tests;

/// <summary>
/// Notes, comments and equations of an ODT: read as the DOCX reader gives them, kept byte for byte
/// in an untouched block and written again in an edited one.
/// </summary>
public class OdtReferenceTests
{
    private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

    private static DocumentModelDto ModelOf(byte[] odt) =>
        JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(OdtReader.Read(odt).Model, JsonOptions.Default), JsonOptions.Default)!;

    private static (byte[] Bytes, SaveResult Result) Save(byte[] odt, DocumentModelDto model) =>
        OdtWriter.Write(odt, JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(model, JsonOptions.Default), JsonOptions.Default)!);

    private static string Part(byte[] zip, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(name)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static List<string> EntryNames(byte[] zip)
    {
        using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        return [.. archive.Entries.Select(entry => entry.FullName)];
    }

    private static List<Node> Blocks(byte[] odt) => OdtReader.Read(odt).Model.Doc.Content!;

    private static string TextOf(Node node) => node.Text ?? string.Concat((node.Content ?? []).Select(TextOf));

    private static string? StringAttr(Node node, string name) =>
        node.Attrs?.GetValueOrDefault(name) is { } value ? value.GetValue<string>() : null;

    private static List<string> Invisible(byte[] odt) => [.. OdtReader.Read(odt).Inventory.Invisible.Select(notice => notice.Key)];

    private static List<string> Lost(SaveResult result) => [.. result.Inventory.Lost.Select(notice => notice.Key)];

    private static string Shape(Node node) =>
        node.Type == "text" ? node.Text! : node.Type + (node.Attrs?.GetValueOrDefault("cid") is { } cid ? ":" + cid.GetValue<string>() : string.Empty);

    private static List<string> Shapes(Node block) => [.. (block.Content ?? []).Select(Shape)];

    private const string TwoNotes =
        """<text:p>Um<text:note text:id="LvNote1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>primeira</text:p></text:note-body></text:note></text:p>""" +
        """<text:p>Dois<text:note text:id="ftn2" text:note-class="footnote"><text:note-citation>2</text:note-citation><text:note-body><text:p>segunda</text:p></text:note-body></text:note></text:p>""";

    [Fact]
    public void FootnotesAndEndnotesBecomeNoteRefsWithTheirBody()
    {
        var odt = Build(
            """<text:p>a<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>nota <text:span text:style-name="T1">forte</text:span></text:p><text:p>outra</text:p></text:note-body></text:note> b""" +
            """<text:note text:id="edn1" text:note-class="endnote"><text:note-citation text:label="*">*</text:note-citation><text:note-body><text:p>fim</text:p></text:note-body></text:note></text:p>""",
            """<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>""");

        var content = Blocks(odt)[0].Content!;
        Assert.Equal(["text", "noteRef", "text", "noteRef"], content.Select(node => node.Type));
        Assert.Equal(" b", content[2].Text);
        var footnote = content[1];
        Assert.Equal(("footnote", "ftn1", null), (StringAttr(footnote, "kind"), StringAttr(footnote, "nid"), StringAttr(footnote, "mark")));
        Assert.Equal(["nota forte", "outra"], footnote.Content!.Select(TextOf));
        Assert.Contains(footnote.Content![0].Content![1].Marks!, mark => mark.Type == "bold");
        var endnote = content[3];
        Assert.Equal(("endnote", "edn1", "*"), (StringAttr(endnote, "kind"), StringAttr(endnote, "nid"), StringAttr(endnote, "mark")));
        Assert.Equal("fim", TextOf(endnote));
        Assert.Empty(Invisible(odt));
    }

    [Fact]
    public void ANoteInsideANoteIsDeclared()
    {
        var odt = Build(
            """<text:p>a<text:note text:note-class="footnote"><text:note-body><text:p>b<text:note text:note-class="endnote"><text:note-body><text:p>c</text:p></text:note-body></text:note></text:p></text:note-body></text:note></text:p>""");

        Assert.Equal("b", TextOf(Blocks(odt)[0].Content![1]));
        Assert.Equal(["sidecar.notice.endnotes"], Invisible(odt));
    }

    [Fact]
    public void NoteNumberingComesFromTheNotesConfiguration()
    {
        var odt = Build(
            "<text:p>a</text:p>",
            styles: LibreOfficeStyles +
                    """<text:notes-configuration text:note-class="footnote" style:num-format="i" text:start-value="2" text:start-numbering-at="page"/>""" +
                    """<text:notes-configuration text:note-class="endnote" style:num-format="A" text:start-value="0" text:start-numbering-at="document"/>""");

        var notes = OdtReader.Read(odt).Model.Notes!;
        Assert.Equal(new NotePrDto("lowerRoman", 3, "eachPage"), notes.FootnotePr);
        Assert.Equal(new NotePrDto("upperLetter", 1), notes.EndnotePr);
    }

    [Fact]
    public void AnUntouchedBlockWithANoteGoesBackAsItWas()
    {
        var odt = Build(TwoNotes);
        var (saved, result) = Save(odt, ModelOf(odt));

        Assert.Equal(Part(odt, "content.xml"), Part(saved, "content.xml"));
        Assert.Equal((2, 0), (result.PreservedBlocks, result.RewrittenBlocks));
    }

    [Fact]
    public void AnEditedNoteFollowsTheKeptOnesNumberAndKeepsItsId()
    {
        var odt = Build(TwoNotes);
        var model = ModelOf(odt);
        model.Doc.Content![1].Content![0].Text = "Dois editado";

        var (saved, result) = Save(odt, model);

        var text = Part(saved, "content.xml");
        Assert.Contains(TwoNotes[..TwoNotes.IndexOf("<text:p>Dois", StringComparison.Ordinal)], text, StringComparison.Ordinal);
        Assert.Matches(
            "<text:p[^>]*>Dois editado<text:note text:id=\"ftn2\" text:note-class=\"footnote\"><text:note-citation>2</text:note-citation><text:note-body><text:p[^>]*>segunda</text:p>",
            text);
        Assert.Equal((1, 1), (result.PreservedBlocks, result.RewrittenBlocks));
        Assert.Empty(Lost(result));
    }

    [Fact]
    public void ANewNoteTakesAnIdNoNoteOfTheOriginalHas()
    {
        var odt = Build(TwoNotes);
        var model = ModelOf(odt);
        var copy = JsonSerializer.Deserialize<Node>(JsonSerializer.Serialize(model.Doc.Content![1].Content![1], DocxJson.Options), DocxJson.Options)!;
        copy.Attrs!.Remove("nid");
        model.Doc.Content.Add(Node.Of("paragraph", new Node { Type = "text", Text = "Três" }, copy));

        var text = Part(Save(odt, model).Bytes, "content.xml");

        Assert.Contains(TwoNotes, text, StringComparison.Ordinal);
        Assert.Matches("Três<text:note text:id=\"LvNote2\"[^>]*><text:note-citation>3</text:note-citation>", text);
    }

    [Fact]
    public void ARewrittenNoteKeepsTheIdAReferenceInAKeptBlockPointsTo()
    {
        var odt = Build(
            """<text:p>Com nota<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>n</text:p></text:note-body></text:note></text:p>""" +
            """<text:p>Veja a nota <text:note-ref text:note-class="footnote" text:reference-format="text" text:ref-name="ftn1">1</text:note-ref></text:p>""");
        var model = ModelOf(odt);
        model.Doc.Content![0].Content![0].Text = "Com nota editada";

        var (saved, result) = Save(odt, model);

        var text = Part(saved, "content.xml");
        Assert.Contains("Com nota editada<text:note text:id=\"ftn1\"", text, StringComparison.Ordinal);
        Assert.Contains("text:ref-name=\"ftn1\">1</text:note-ref>", text, StringComparison.Ordinal);
        Assert.Equal((1, 1), (result.PreservedBlocks, result.RewrittenBlocks));
    }

    [Fact]
    public void EditingOnlyANoteBodyRewritesItsBlock()
    {
        var odt = Build(TwoNotes);
        var model = ModelOf(odt);
        model.Doc.Content![0].Content![1].Content![0].Content![0].Text = "primeira revista";

        var (saved, result) = Save(odt, model);

        Assert.Equal("primeira revista", TextOf(Blocks(saved)[0].Content![1]));
        Assert.Equal((1, 1), (result.PreservedBlocks, result.RewrittenBlocks));
    }

    private const string Threads =
        """<text:p>a<office:annotation office:name="__Annotation__7_1" loext:resolved="true"><dc:creator>Bia</dc:creator><dc:date>2026-03-01T10:00:00</dc:date><meta:creator-initials>BS</meta:creator-initials><text:p>Confira</text:p><text:p>linha<text:s/> 2</text:p></office:annotation>texto<office:annotation-end office:name="__Annotation__7_1"/>""" +
        """<office:annotation office:name="__Annotation__8_1" loext:parent-name="__Annotation__7_1" loext:resolved="true"><dc:creator>Ana</dc:creator><text:p>Feito</text:p></office:annotation> b<office:annotation><dc:creator>Cid</dc:creator><text:p>ponto</text:p></office:annotation></text:p>""";

    [Fact]
    public void CommentsBecomeAnchorsAndTheLibraryWithRepliesAndResolution()
    {
        var odt = Build(Threads);
        var model = OdtReader.Read(odt).Model;

        Assert.Equal(["a", "commentStart:0", "texto", "commentEnd:0", " b", "commentEnd:2"], Shapes(model.Doc.Content![0]));
        Assert.Equal(
            [
                new CommentDto("0", "Bia", "2026-03-01T10:00:00", ["Confira", "linha  2"], true, Initials: "BS"),
                new CommentDto("1", "Ana", string.Empty, ["Feito"], true, ParentId: "0"),
                new CommentDto("2", "Cid", string.Empty, ["ponto"], false),
            ],
            model.Comments!,
            CommentComparer.Instance);
        Assert.Empty(Invisible(odt));
    }

    [Fact]
    public void AnUntouchedBlockWithCommentsGoesBackAsItWas()
    {
        var odt = Build(Threads);
        var (saved, result) = Save(odt, ModelOf(odt));

        Assert.Equal(Part(odt, "content.xml"), Part(saved, "content.xml"));
        Assert.Equal(0, result.RewrittenBlocks);
    }

    [Fact]
    public void AnEditedBlockWritesItsCommentsAgainWithTheirNames()
    {
        var odt = Build(Threads);
        var model = ModelOf(odt);
        model.Doc.Content![0].Content![0].Text = "A";

        var (saved, result) = Save(odt, model);

        var text = Part(saved, "content.xml");
        Assert.Contains("<office:annotation office:name=\"__Annotation__7_1\" loext:resolved=\"true\"><dc:creator>Bia</dc:creator>", text, StringComparison.Ordinal);
        Assert.Contains("<office:annotation-end office:name=\"__Annotation__7_1\"/>", text, StringComparison.Ordinal);
        Assert.Contains("loext:parent-name=\"__Annotation__7_1\"", text, StringComparison.Ordinal);
        var reread = OdtReader.Read(saved).Model;
        Assert.Equal(["A", "commentStart:0", "texto", "commentEnd:0", " b", "commentEnd:2"], Shapes(reread.Doc.Content![0]));
        Assert.Equal(OdtReader.Read(odt).Model.Comments!, reread.Comments!, CommentComparer.Instance);
        Assert.Empty(Lost(result));
    }

    [Fact]
    public void ResolvingACommentRewritesTheBlockThatHoldsIt()
    {
        var odt = Build(Threads + "<text:p>outro</text:p>");
        var model = ModelOf(odt);
        model = model with { Comments = [.. model.Comments!.Select(comment => comment.Id == "2" ? comment with { Done = true } : comment)] };

        var (saved, result) = Save(odt, model);

        Assert.True(OdtReader.Read(saved).Model.Comments![2].Done);
        Assert.Equal((1, 1), (result.PreservedBlocks, result.RewrittenBlocks));
    }

    [Fact]
    public void AReplyWrittenBeforeItsCommentStillAnswersIt()
    {
        var odt = Build(
            """<text:p>a<office:annotation office:name="B" loext:parent-name="A"><dc:creator>Ana</dc:creator><text:p>resposta</text:p></office:annotation>""" +
            """<office:annotation office:name="A"><dc:creator>Bia</dc:creator><text:p>pergunta</text:p></office:annotation>x<office:annotation-end office:name="A"/><office:annotation-end office:name="B"/></text:p>""");

        var model = OdtReader.Read(odt).Model;
        Assert.Equal(["a", "commentStart:1", "x", "commentEnd:1"], Shapes(model.Doc.Content![0]));
        Assert.Equal([("0", "1"), ("1", null)], model.Comments!.Select(comment => (comment.Id, comment.ParentId)));

        var edited = ModelOf(odt);
        edited.Doc.Content![0].Content![0].Text = "b";
        var text = Part(Save(odt, edited).Bytes, "content.xml");
        Assert.Contains("<office:annotation office:name=\"A\"><dc:creator>Bia</dc:creator><text:p>pergunta</text:p></office:annotation><office:annotation loext:parent-name=\"A\"><dc:creator>Ana</dc:creator>", text, StringComparison.Ordinal);
    }

    private const string CrossingRange =
        """<text:p>um<office:annotation office:name="Faixa"><dc:creator>Bia</dc:creator><text:p>c</text:p></office:annotation> começo</text:p>""" +
        """<text:p>dois fim<office:annotation-end office:name="Faixa"/> resto</text:p>""";

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void ARangeAcrossBlocksStillClosesWhenOnlyOneEndIsRewritten(int edited)
    {
        var odt = Build(CrossingRange);
        var model = ModelOf(odt);
        model.Doc.Content![edited].Content![0].Text += "!";

        var (saved, result) = Save(odt, model);

        var text = Part(saved, "content.xml");
        Assert.Contains("<office:annotation office:name=\"Faixa\">", text, StringComparison.Ordinal);
        Assert.Contains("<office:annotation-end office:name=\"Faixa\"/>", text, StringComparison.Ordinal);
        var blocks = Blocks(saved);
        Assert.Contains("commentStart:0", Shapes(blocks[0]));
        Assert.Contains("commentEnd:0", Shapes(blocks[1]));
        Assert.Equal((1, 1), (result.PreservedBlocks, result.RewrittenBlocks));
    }

    [Fact]
    public void ANewCommentNeverTakesTheNameOfAKeptOne()
    {
        var odt = Build(
            """<text:p>a<office:annotation office:name="__Annotation__1"><dc:creator>Bia</dc:creator><text:p>c</text:p></office:annotation>b<office:annotation-end office:name="__Annotation__1"/></text:p><text:p>novo</text:p>""");
        var model = ModelOf(odt);
        model.Doc.Content![1].Content =
        [
            Node.Of("commentStart").With("cid", "1"),
            new Node { Type = "text", Text = "novo" },
            Node.Of("commentEnd").With("cid", "1"),
        ];
        model = model with { Comments = [.. model.Comments!, new CommentDto("1", "Ana", string.Empty, ["nova"], false)] };

        var (saved, _) = Save(odt, model);

        var text = Part(saved, "content.xml");
        Assert.Contains("<office:annotation office:name=\"__Annotation__1_1\"><dc:creator>Ana</dc:creator>", text, StringComparison.Ordinal);
        Assert.Contains("<office:annotation-end office:name=\"__Annotation__1_1\"/>", text, StringComparison.Ordinal);
        Assert.Equal(["Bia", "Ana"], OdtReader.Read(saved).Model.Comments!.Select(comment => comment.Author));
    }

    [Fact]
    public void APointCommentFromTheEditorIsWrittenWhereItsEndIs()
    {
        var odt = Build("<text:p>a</text:p>");
        var model = ModelOf(odt);
        model.Doc.Content![0].Content!.Add(Node.Of("commentEnd").With("cid", "0"));
        model = model with { Comments = [new CommentDto("0", "Ana", string.Empty, ["aqui"], false)] };

        var (saved, result) = Save(odt, model);

        Assert.Matches("<text:p[^>]*>a<office:annotation><dc:creator>Ana</dc:creator><text:p>aqui</text:p></office:annotation></text:p>", Part(saved, "content.xml"));
        Assert.Equal(["a", "commentEnd:0"], Shapes(Blocks(saved)[0]));
        Assert.Empty(Lost(result));
    }

    [Fact]
    public void ACommentWithMoreThanTextIsLostOnlyWhenItsBlockIsRewritten()
    {
        var odt = Build(
            """<text:p>a<office:annotation><dc:creator>Bia</dc:creator><text:p>veja <text:a xlink:href="https://example.com">aqui</text:a></text:p></office:annotation></text:p>""");
        Assert.True(OdtReader.Read(odt).Model.Comments![0].Rich);
        Assert.Equal(["sidecar.notice.comments"], Invisible(odt));

        var model = ModelOf(odt);
        model.Doc.Content![0].Content![0].Text = "b";
        Assert.Equal(["sidecar.notice.comments"], Lost(Save(odt, model).Result));
    }

    private static byte[] Formula(string math) => Utf8("""<?xml version="1.0" encoding="UTF-8"?>""" + math);

    private static byte[] WithFormula(string body, string math) =>
        Build(body, files: new Dictionary<string, byte[]> { ["Object 1/content.xml"] = Formula(math) });

    private const string InlineFormula =
        """<text:p>Seja <draw:frame draw:name="Objeto1" text:anchor-type="as-char" svg:width="5mm" svg:height="4mm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><svg:desc>x^2</svg:desc></draw:frame> fim</text:p>""";

    private const string OldLibreOfficeFormula =
        """<math:math xmlns:math="http://www.w3.org/1998/Math/MathML" display="block"><math:semantics><math:msup><math:mi>x</math:mi><math:mn>2</math:mn></math:msup><math:annotation encoding="StarMath 5.0">x^2</math:annotation></math:semantics></math:math>""";

    [Fact]
    public void AFormulaObjectBecomesAnEquationThroughTheFilter()
    {
        var odt = WithFormula(InlineFormula, OldLibreOfficeFormula);

        var content = Blocks(odt)[0].Content!;
        Assert.Equal(["text", "math", "text"], content.Select(node => node.Type));
        var math = content[1];
        Assert.Equal(
            $"""<math display="block" xmlns="{MathMlNamespace}"><semantics><msup><mi>x</mi><mn>2</mn></msup><annotation encoding="StarMath 5.0">x^2</annotation></semantics></math>""",
            StringAttr(math, "mathml"));
        Assert.Equal("x^2", StringAttr(math, "latex"));
        Assert.False(math.Attrs!["display"]!.GetValue<bool>());
        Assert.Empty(Invisible(odt));
    }

    [Fact]
    public void AFormulaAloneInItsParagraphIsDisplayed()
    {
        var odt = WithFormula(
            """<text:p><draw:frame text:anchor-type="as-char"><draw:object xlink:href="./Object 1"/></draw:frame></text:p>""",
            $"""<math xmlns="{MathMlNamespace}" display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>""");

        Assert.True(Blocks(odt)[0].Content![0].Attrs!["display"]!.GetValue<bool>());
    }

    [Fact]
    public void WhatTheFilterDropsFromAFormulaIsDeclared()
    {
        var odt = WithFormula(InlineFormula, $"""<math xmlns="{MathMlNamespace}"><menclose><mi>x</mi></menclose></math>""");

        Assert.Equal($"""<math xmlns="{MathMlNamespace}" />""", StringAttr(Blocks(odt)[0].Content![1], "mathml"));
        Assert.Equal(["sidecar.notice.equations"], Invisible(odt));
    }

    [Fact]
    public void AnUntouchedFormulaKeepsItsObject()
    {
        var odt = WithFormula(InlineFormula, OldLibreOfficeFormula);
        var (saved, result) = Save(odt, ModelOf(odt));

        Assert.Equal(Part(odt, "content.xml"), Part(saved, "content.xml"));
        Assert.Equal(OldLibreOfficeFormula, Part(saved, "Object 1/content.xml")[38..]);
        Assert.Equal(0, result.RewrittenBlocks);
    }

    [Fact]
    public void TheLatexTheEditorDerivesIsNotAnEditOfTheFormula()
    {
        var odt = WithFormula(InlineFormula.Replace("<svg:desc>x^2</svg:desc>", string.Empty, StringComparison.Ordinal), OldLibreOfficeFormula);
        var model = ModelOf(odt);
        var math = model.Doc.Content![0].Content![1];
        math.With("latex", "x^{2}").With("lossy", null).With("editable", null).With("jc", null);

        var (saved, result) = Save(odt, model);

        Assert.Equal(Part(odt, "content.xml"), Part(saved, "content.xml"));
        Assert.Equal(0, result.RewrittenBlocks);
    }

    [Fact]
    public void AnEditedBlockWritesItsFormulaAsANewObjectThatReadsTheSame()
    {
        var odt = WithFormula(InlineFormula, OldLibreOfficeFormula);
        var model = ModelOf(odt);
        model.Doc.Content![0].Content![0].Text = "Então ";

        var (saved, result) = Save(odt, model);

        Assert.DoesNotContain("Object 1/content.xml", EntryNames(saved));
        var before = Blocks(odt)[0].Content![1];
        var after = Blocks(saved)[0].Content![1];
        Assert.Equal(("math", StringAttr(before, "latex")), (after.Type, StringAttr(after, "latex")));
        Assert.Contains("<mi>x</mi><mn>2</mn>", StringAttr(after, "mathml"), StringComparison.Ordinal);
        Assert.Empty(Lost(result));
    }

    /// <summary>Records compare their lists by reference: the paragraphs are compared by content.</summary>
    private sealed class CommentComparer : IEqualityComparer<CommentDto>
    {
        public static readonly CommentComparer Instance = new();

        private static readonly List<string> None = [];

        public bool Equals(CommentDto? x, CommentDto? y) =>
            x is not null && y is not null && x.Paragraphs.SequenceEqual(y.Paragraphs) && x with { Paragraphs = None } == y with { Paragraphs = None };

        public int GetHashCode(CommentDto obj) => obj.Id.GetHashCode(StringComparison.Ordinal);
    }
}

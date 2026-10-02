using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// <summary>
/// Notas de rodapé e de fim (M11, fase 1): lidas, preservadas, devolvidas ao arquivo.
/// </summary>
/// <remarks>
/// A referência virou o nó `noteRef`, com o corpo da nota dentro. A impressão
/// digital do parágrafo vê só para onde ela aponta — editar a nota não reescreve o
/// parágrafo —, e a parte das notas só é tocada quando alguma nota mudou.
/// </remarks>
public class NotesTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Styles =
        """<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>""" +
        """<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>""" +
        """<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>""" +
        """<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/></w:style>""" +
        """<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>""";

    private const string Body =
        """<w:p><w:r><w:t>Intocado.</w:t></w:r></w:p>""" +
        """<w:p><w:r><w:t>Texto com nota</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>""" +
        """<w:p><w:r><w:t>Marca própria</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:customMarkFollows="1" w:id="2"/><w:t>*</w:t></w:r></w:p>""" +
        """<w:p><w:r><w:t>Com nota de fim</w:t></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/></w:r></w:p>""";

    private static string Separators(string kind) =>
        $"""<w:{kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}>""" +
        $"""<w:{kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>""";

    private const string Footnotes =
        """<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Fonte: ata anterior.</w:t></w:r></w:p>""" +
        """<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t>Segundo parágrafo da nota.</w:t></w:r></w:p></w:footnote>""" +
        """<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t xml:space="preserve"> Nota de marca própria.</w:t></w:r></w:p></w:footnote>""";

    private const string Endnotes =
        """<w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Ao fim.</w:t></w:r></w:p></w:endnote>""";

    private const string Settings =
        """<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>""" +
        """<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>""";

    /// <summary>Um documento com duas notas de rodapé (uma de marca própria) e uma de fim, como o Word grava.</summary>
    internal static byte[] WithNotes()
    {
        var ns = $"xmlns:w=\"{W}\"";
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var part = document.AddMainDocumentPart();
            var styles = part.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles($"<w:styles {ns}>{Styles}</w:styles>");
            styles.Styles.Save();

            var settings = part.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings(
                $"<w:settings {ns}>{Settings}</w:settings>");
            settings.Settings.Save();

            var footnotes = part.AddNewPart<FootnotesPart>();
            footnotes.Footnotes = new DocumentFormat.OpenXml.Wordprocessing.Footnotes(
                $"<w:footnotes {ns}>{Separators("footnote")}{Footnotes}</w:footnotes>");
            footnotes.Footnotes.Save();

            var endnotes = part.AddNewPart<EndnotesPart>();
            endnotes.Endnotes = new DocumentFormat.OpenXml.Wordprocessing.Endnotes(
                $"<w:endnotes {ns}>{Separators("endnote")}{Endnotes}</w:endnotes>");
            endnotes.Endnotes.Save();

            part.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                $"<w:document {ns}><w:body>{Body}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
                "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>");
            part.Document.Save();
        }

        return buffer.ToArray();
    }

    private static List<Node> NoteRefs(DocumentModelDto model) =>
        Walk(model.Doc).Where(node => node.Type == "noteRef").ToList();

    private static Node NoteRef(DocumentModelDto model, string kind, string nid) =>
        NoteRefs(model).Single(node => Attr(node, "kind") == kind && Attr(node, "nid") == nid);

    private static string? Attr(Node node, string name) =>
        node.Attrs?.GetValueOrDefault(name)?.GetValue<string>();

    private static Node Paragraph(Node doc, string startsWith) =>
        doc.Content!.First(block => string.Concat(Walk(block)
            .TakeWhile(node => node.Type != "noteRef")
            .Where(node => node.Type == "text").Select(node => node.Text)).StartsWith(startsWith, StringComparison.Ordinal));

    /// <summary>O `w:footnote`/`w:endnote` de id dado, no XML gravado.</summary>
    private static string NoteXml(string xml, string kind, string id) =>
        Regex.Match(xml, $"""<w:{kind} [^>]*w:id="{id}"[^>]*>.*?</w:{kind}>""", RegexOptions.Singleline).Value;

    [Fact]
    public void LeAReferenciaComOCorpoDaNotaDentro()
    {
        var result = DocxReader.Read(WithNotes());
        var model = result.Model;

        var first = NoteRef(model, "footnote", "1");
        Assert.Equal(2, first.Content!.Count);
        Assert.Equal(" Fonte: ata anterior.", string.Concat(Walk(first.Content[0]).Select(node => node.Text)));
        Assert.Equal("fn:1/p1", Attr(first.Content[0], "oid"));
        Assert.Equal("fn:1/p2", Attr(first.Content[1], "oid"));

        var custom = NoteRef(model, "footnote", "2");
        Assert.Equal("*", Attr(custom, "mark"));
        Assert.DoesNotContain(Walk(model.Doc), node => node.Type == "text" && node.Text == "*");

        Assert.Equal(" Ao fim.", string.Concat(Walk(NoteRef(model, "endnote", "1")).Select(node => node.Text)));

        // A numeração do documento, fora dos nós.
        Assert.Equal("lowerRoman", model.Notes?.FootnotePr?.NumFmt);
        Assert.Equal(3, model.Notes?.FootnotePr?.Start);

        // A nota deixou de ser invisível — e de travar o documento.
        Assert.DoesNotContain(Inventory.Footnotes, result.Inventory.Invisible);
        Assert.DoesNotContain(Inventory.Endnotes, result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void AbrirESalvarNaoTocaNadaDasNotas()
    {
        var original = WithNotes();
        var (bytes, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);

        var before = PartsOf(original);
        var after = PartsOf(bytes);
        foreach (var path in new[] { "word/footnotes.xml", "word/endnotes.xml", "word/settings.xml", "word/styles.xml" })
        {
            Assert.Equal(before[path], after[path]);
        }
    }

    [Fact]
    public void EditarUmaNotaReescreveSoOParagrafoDelaENaoOQueAReferencia()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        var note = NoteRef(model, "footnote", "1");
        Walk(note.Content![0]).First(node => node.Type == "text").Text = " Fonte: ata corrigida.";

        var (bytes, result) = Save(original, model);

        // A nota mudou; o parágrafo que a referencia, não.
        Assert.Equal(1, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);

        var before = XmlOf(original, "word/footnotes.xml");
        var after = XmlOf(bytes, "word/footnotes.xml");
        // A parte foi gravada pelo SDK, que fecha o elemento vazio com " />": é a
        // única diferença de escrita na nota que ninguém tocou.
        Assert.Equal(NoteXml(before, "footnote", "2"), NoteXml(after, "footnote", "2").Replace(" />", "/>"));
        Assert.Equal(NoteXml(before, "footnote", "-1"), NoteXml(after, "footnote", "-1").Replace(" />", "/>"));

        var edited = NoteXml(after, "footnote", "1");
        Assert.Contains("Fonte: ata corrigida.", edited);
        Assert.Contains("<w:footnoteRef />", edited.Replace("<w:footnoteRef/>", "<w:footnoteRef />"));
        Assert.Contains("Segundo parágrafo da nota.", edited);
        Assert.Contains("FootnoteText", edited);

        // As outras partes voltam byte a byte.
        Assert.Equal(PartsOf(original)["word/endnotes.xml"], PartsOf(bytes)["word/endnotes.xml"]);

        var reread = NoteRef(Open(bytes), "footnote", "1");
        Assert.Equal(" Fonte: ata corrigida.", string.Concat(Walk(reread.Content![0]).Select(node => node.Text)));
    }

    [Fact]
    public void RevisaoNovaNaNotaGanhaIdUnicoNoPacote()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        // Duas inserções novas, sem `rid`: uma na nota de rodapé, outra no corpo.
        var paragraph = NoteRef(model, "footnote", "1").Content![0];
        paragraph.Content!.Add(new Node
        {
            Type = "text",
            Text = " acrescentado",
            Marks = [new Mark { Type = Revisions.Insertion, Attrs = new() { ["author"] = "Ana", ["date"] = "2026-01-01T00:00:00Z" } }],
        });
        Paragraph(model.Doc, "Intocado.").Content!.Add(new Node
        {
            Type = "text",
            Text = " novo",
            Marks = [new Mark { Type = Revisions.Insertion, Attrs = new() { ["author"] = "Ana", ["date"] = "2026-01-01T00:00:00Z" } }],
        });

        var (bytes, _) = Save(original, model);

        var ids = new[] { "word/document.xml", "word/footnotes.xml" }
            .SelectMany(path => Regex.Matches(XmlOf(bytes, path), @"<w:ins\b[^>]*>").Select(tag => tag.Value))
            .Select(tag => Regex.Match(tag, "w:id=\"([^\"]+)\"").Groups[1].Value)
            .ToList();
        Assert.Equal(2, ids.Count);
        Assert.All(ids, id => Assert.NotEqual("", id));
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void ApagarAReferenciaTiraANota()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        var host = Paragraph(model.Doc, "Texto com nota");
        host.Content = [.. host.Content!.Where(node => node.Type != "noteRef")];

        var (bytes, result) = Save(original, model);
        Assert.Empty(result.Inventory.Lost);

        var after = XmlOf(bytes, "word/footnotes.xml");
        Assert.Empty(NoteXml(after, "footnote", "1"));
        Assert.NotEmpty(NoteXml(after, "footnote", "2"));
        Assert.Contains("w:separator", after);
        Assert.Contains("w:continuationSeparator", after);
    }

    [Fact]
    public void ReferenciaMovidaLevaANotaParaOMesmoLugarNaParte()
    {
        // O LibreOffice casa nota e referência pela ordem na parte: movida a
        // referência (recortar e colar), a nota muda de lugar com ela.
        var original = WithNotes();
        var model = Clone(Open(original));
        var first = NoteRef(model, "footnote", "1");
        var second = NoteRef(model, "footnote", "2");
        var from = Walk(model.Doc).First(node => node.Content?.Contains(first) == true);
        var to = Walk(model.Doc).First(node => node.Content?.Contains(second) == true);
        from.Content!.Remove(first);
        to.Content!.Add(first);

        var (bytes, result) = Save(original, model);
        Assert.Empty(result.Inventory.Lost);

        // Renumeradas pela ordem do texto, como o Word grava: a que passou a vir
        // primeiro é a 1, e cada corpo segue a sua referência.
        var after = XmlOf(bytes, "word/footnotes.xml");
        Assert.Contains("Nota de marca própria", NoteXml(after, "footnote", "1"));
        Assert.Contains("Fonte: ata anterior", NoteXml(after, "footnote", "2"));
        Assert.Equal(["1", "2"], NoteRefs(Open(bytes)).Where(node => Attr(node, "kind") == "footnote")
            .Select(node => Attr(node, "nid")!));
    }

    [Fact]
    public void ReferenciaRepetidaGanhaNotaPropria()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        var copy = Clone(model with { }).Doc;
        var pasted = NoteRefs(new DocumentModelDto(model.Page, copy)).First(node => Attr(node, "nid") == "1");
        var target = Paragraph(model.Doc, "Intocado.");
        target.Content = [.. target.Content!, pasted];

        var (bytes, _) = Save(original, model);

        // A cópia ganha nota própria, com o mesmo texto; os ids seguem a ordem do
        // texto (`NotesWriter.InTextOrder`), e a cópia, no primeiro parágrafo, vem antes.
        var after = XmlOf(bytes, "word/footnotes.xml");
        var copies = new[] { "1", "2", "3" }.Select(id => NoteXml(after, "footnote", id))
            .Where(note => note.Contains("Fonte: ata anterior.", StringComparison.Ordinal)).ToList();
        Assert.Equal(2, copies.Count);
        Assert.All(copies, note => Assert.Contains("footnoteRef", note));

        var reread = Open(bytes);
        Assert.Equal(["1", "2", "3"], NoteRefs(reread).Where(node => Attr(node, "kind") == "footnote")
            .Select(node => Attr(node, "nid")!).Order(StringComparer.Ordinal));
    }

    [Fact]
    public void ParagrafoReescritoDevolveAReferenciaComoEra()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        Assert.True(EditFirstTextContaining(model, "Texto com nota", "Texto editado"));
        Assert.True(EditFirstTextContaining(model, "Marca própria", "Marca editada"));

        var (bytes, result) = Save(original, model);
        Assert.Equal(2, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);

        var xml = XmlOf(bytes);
        Assert.Matches("""<w:r><w:rPr><w:rStyle w:val="FootnoteReference" ?/></w:rPr><w:footnoteReference w:id="1" ?/></w:r>""", xml);
        Assert.Matches("""<w:footnoteReference w:customMarkFollows="(1|true)" w:id="2" ?/><w:t[^>]*>\*</w:t>""", xml);

        // As notas não mudaram: a parte volta byte a byte.
        Assert.Equal(PartsOf(original)["word/footnotes.xml"], PartsOf(bytes)["word/footnotes.xml"]);

        var reread = Open(bytes);
        Assert.Equal("*", Attr(NoteRef(reread, "footnote", "2"), "mark"));
    }

    [Fact]
    public void NotaDeFimEditadaNaoTocaAsDeRodape()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        Walk(NoteRef(model, "endnote", "1")).First(node => node.Type == "text").Text = " Ao fim, revisto.";

        var (bytes, _) = Save(original, model);

        Assert.Contains("Ao fim, revisto.", XmlOf(bytes, "word/endnotes.xml"));
        Assert.Contains("endnoteRef", XmlOf(bytes, "word/endnotes.xml"));
        Assert.Equal(PartsOf(original)["word/footnotes.xml"], PartsOf(bytes)["word/footnotes.xml"]);
    }

    [Fact]
    public void NotaNovaNumDocumentoSemNotasCriaAParte()
    {
        var original = Fixtures.Simple();
        var model = Clone(Open(original));
        var host = model.Doc.Content!.First(block => block.Type == "paragraph");
        var body = Node.Of("paragraph", new Node { Type = "text", Text = "Nota nova." });
        var reference = Node.Of("noteRef", body).With("kind", "footnote").With("nid", null);
        host.Content = [.. host.Content ?? [], reference];

        var (bytes, result) = Save(original, model);
        Assert.Empty(result.Inventory.Lost);

        var parts = PartsOf(bytes);
        Assert.True(parts.ContainsKey("word/footnotes.xml"));
        var notes = XmlOf(bytes, "word/footnotes.xml");
        Assert.Contains("w:separator", notes);
        Assert.Contains("w:continuationSeparator", notes);
        Assert.Contains("Nota nova.", NoteXml(notes, "footnote", "1"));
        Assert.Contains("footnoteRef", NoteXml(notes, "footnote", "1"));
        Assert.Contains("footnotes+xml", XmlOf(bytes, "[Content_Types].xml"));
        Assert.Matches("""<w:footnotePr><w:footnote w:id="-1" ?/><w:footnote w:id="0" ?/></w:footnotePr>""", XmlOf(bytes, "word/settings.xml"));
        Assert.Contains("<w:vertAlign w:val=\"superscript\" />", XmlOf(bytes).Replace("\"/>", "\" />"));

        var reread = NoteRefs(Open(bytes)).Single();
        Assert.Equal("1", Attr(reread, "nid"));
        Assert.Equal("Nota nova.", string.Concat(Walk(reread).Select(node => node.Text)));
    }

    [Fact]
    public void RascunhoDeAntesDasNotasDeclaraAPerdaENaoMexeNasPartes()
    {
        var original = WithNotes();
        // O modelo como o leitor de antes das notas o dava: sem `noteRef`.
        List<Node> content;
        using (var stream = new MemoryStream(original))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var part = document.MainDocumentPart!;
            (content, _) = new BodyReader(part, new Inventory(), notes: false).Read(part.Document!.Body!);
        }

        var doc = Node.Of("doc");
        doc.Content = content;
        var model = Clone(new DocumentModelDto(Open(original).Page, doc));
        Assert.DoesNotContain(Walk(model.Doc), node => node.Type == "noteRef");

        var untouched = Save(original, model with { BeforeNotes = true });
        Assert.Equal(0, untouched.Result.RewrittenBlocks);

        Assert.True(EditFirstTextContaining(model, "Texto com nota", "Texto editado"));
        var (bytes, result) = Save(original, model with { BeforeNotes = true });
        Assert.Contains("nota de rodapé num parágrafo que você editou", result.Inventory.Lost);
        Assert.Equal(PartsOf(original)["word/footnotes.xml"], PartsOf(bytes)["word/footnotes.xml"]);
    }

    [Fact]
    public void NumeracaoDoModeloQueOPacoteNaoTemEGravada()
    {
        // O `.sdoc` reaberto guarda a numeração no modelo; o pacote em que ele é
        // gravado como `.docx` não a tem — e ela não pode se perder.
        var original = WithNotes();
        var model = Open(original);
        Assert.Equal(3, model.Notes?.FootnotePr?.Start);

        byte[] target;
        using (var stream = new MemoryStream())
        {
            stream.Write(original);
            using (var document = WordprocessingDocument.Open(stream, true))
            {
                var settings = document.MainDocumentPart!.DocumentSettingsPart!.Settings!;
                foreach (var child in settings.Descendants().Where(child => child.LocalName is "numFmt" or "numStart").ToList())
                    child.Remove();
                settings.Save();
            }

            target = stream.ToArray();
        }

        Assert.Null(Open(target).Notes?.FootnotePr);

        var (bytes, _) = Save(target, model);
        Assert.Matches(
            """<w:footnotePr><w:numFmt w:val="lowerRoman" ?/><w:numStart w:val="3" ?/><w:footnote w:id="-1" ?/><w:footnote w:id="0" ?/></w:footnotePr>""",
            XmlOf(bytes, "word/settings.xml"));
        Assert.Equal(model.Notes, Open(bytes).Notes);
        AssertSchema(bytes);
    }

    [Fact]
    public void NumeracaoNovaDaNotaDeFimCriaODeclaradoNoSettings()
    {
        var original = WithNotes();
        var model = Open(original) with
        {
            Notes = new NotesDto(new NotePrDto("lowerRoman", 3), new NotePrDto("upperLetter", 2)),
        };

        var (bytes, _) = Save(original, model);
        var settings = XmlOf(bytes, "word/settings.xml");
        Assert.Matches(
            """<w:endnotePr><w:numFmt w:val="upperLetter" ?/><w:numStart w:val="2" ?/><w:endnote w:id="-1" ?/>""",
            settings);
        Assert.Equal(model.Notes, Open(bytes).Notes);
    }

    [Fact]
    public void NumeracaoIgualNaoTocaOSettings()
    {
        var original = WithNotes();
        var (bytes, _) = Save(original, Open(original));
        Assert.Equal(PartsOf(original)["word/settings.xml"], PartsOf(bytes)["word/settings.xml"]);
    }

    // --- fase 3: comentários nas notas e NOTEREF --------------------------------

    private static Node NodeOf(string type, params (string Name, string Value)[] attrs)
    {
        var node = new Node { Type = type, Attrs = [] };
        foreach (var (name, value) in attrs) node.Attrs[name] = System.Text.Json.Nodes.JsonValue.Create(value);
        return node;
    }

    private static List<string> CommentMarkers(string xml) =>
        [.. Regex.Matches(xml, "<w:(commentRangeStart|commentRangeEnd|commentReference) w:id=\"(\\d+)\"")
            .Select(match => $"{match.Groups[1].Value}:{match.Groups[2].Value}")];

    /// <summary>
    /// O validador do SDK, menos o falso positivo da âncora de comentário numa nota:
    /// ele procura `comments.xml` entre as relações da parte das notas, e o Word não
    /// a relaciona ali — o comentário mora na parte do documento.
    /// </summary>
    private static void AssertSchemaButCommentsInNotes(byte[] docx)
    {
        using var stream = new MemoryStream(docx.ToArray());
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2021)
            .Validate(document)
            .Where(error => !(error.Part is FootnotesPart or EndnotesPart &&
                              error.Description.Contains("'WordprocessingCommentsPart'", StringComparison.Ordinal)))
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();
        Assert.Empty(errors);
    }

    /// <summary>O documento de notas com um comentário novo no trecho da primeira nota de rodapé.</summary>
    private static (byte[] Original, byte[] Saved) WithCommentInNote()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        var paragraph = NoteRef(model, "footnote", "1").Content![0];
        paragraph.Content!.Insert(0, NodeOf("commentStart", ("cid", "0")));
        paragraph.Content.Add(NodeOf("commentEnd", ("cid", "0")));
        model = model with { Comments = [new CommentDto("0", "Vinícius Storch", "2026-10-02T12:00:00Z", ["Conferir a ata."], false)] };
        return (original, DocxWriter.Write(original, model).Bytes);
    }

    [Fact]
    public void ComentarioNovoNaNotaVaiParaANotaEVoltaNaLeitura()
    {
        var (_, saved) = WithCommentInNote();

        // As pontas e a referência na nota, e não no corpo.
        Assert.Equal(["commentRangeStart:0", "commentRangeEnd:0", "commentReference:0"],
            CommentMarkers(NoteXml(XmlOf(saved, "word/footnotes.xml"), "footnote", "1")));
        Assert.Empty(CommentMarkers(XmlOf(saved)));
        Assert.Contains("Conferir a ata.", XmlOf(saved, "word/comments.xml"), StringComparison.Ordinal);
        AssertSchemaButCommentsInNotes(saved);

        var reopened = Open(saved);
        Assert.Equal("0", Assert.Single(reopened.Comments!).Id);
        var anchors = Walk(NoteRef(reopened, "footnote", "1"))
            .Where(node => node.Type is "commentStart" or "commentEnd")
            .Select(node => $"{node.Type}:{Attr(node, "cid")}");
        Assert.Equal(["commentStart:0", "commentEnd:0"], anchors);
    }

    [Fact]
    public void ComentarioExcluidoDaNotaSaiDaNotaEDosComentarios()
    {
        var (_, saved) = WithCommentInNote();
        var model = Clone(Open(saved));
        foreach (var node in Walk(NoteRef(model, "footnote", "1")))
            node.Content?.RemoveAll(child => child.Type is "commentStart" or "commentEnd");
        model = model with { Comments = [] };

        var (bytes, result) = Save(saved, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Empty(CommentMarkers(XmlOf(bytes, "word/footnotes.xml")));
        Assert.Empty(Open(bytes).Comments ?? []);
    }

    [Fact]
    public void ReferenciaCruzadaANotaGravaONoterefEOMarcadorEmVoltaDaReferencia()
    {
        var original = WithNotes();
        var model = Clone(Open(original));
        var host = Paragraph(model.Doc, "Texto com nota");
        var at = host.Content!.FindIndex(node => node.Type == "noteRef");
        host.Content.Insert(at + 1, NodeOf("bookmarkEnd", ("bid", "40")));
        host.Content.Insert(at, NodeOf("bookmarkStart", ("name", "_Ref900"), ("bid", "40")));
        Paragraph(model.Doc, "Intocado.").Content!.Add(NodeOf("field", ("instr", " NOTEREF _Ref900 \\h "), ("result", "iii")));

        var (bytes, _) = Save(original, model);
        var xml = XmlOf(bytes);

        Assert.Matches(new Regex("<w:bookmarkStart [^>]*w:name=\"_Ref900\"[^>]*/>.*<w:footnoteReference w:id=\"1\"[^>]*/>.*<w:bookmarkEnd ", RegexOptions.Singleline), xml);
        Assert.Contains("NOTEREF _Ref900 \\h", xml, StringComparison.Ordinal);
        AssertSchema(bytes);

        var reopened = Open(bytes);
        var field = Walk(reopened.Doc).Single(node => node.Type == "field");
        Assert.Equal(" NOTEREF _Ref900 \\h ", Attr(field, "instr"));
        Assert.Equal("iii", Attr(field, "result"));
        Assert.Contains(Walk(reopened.Doc), node => node.Type == "bookmarkStart" && Attr(node, "name") == "_Ref900");
    }
}

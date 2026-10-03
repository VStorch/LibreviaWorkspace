using System.Text.RegularExpressions;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// <summary>
/// Comentários: lidos, mostrados e devolvidos ao arquivo.
/// </summary>
/// <remarks>
/// A âncora é um par de nós no parágrafo, e o corpo do comentário mora fora dos
/// nós. Editar o parágrafo comentado não perde nada — e por isso o comentário
/// não trava o documento.
/// </remarks>
public class CommentsTests
{
    private static List<string> Anchors(DocumentModelDto model) =>
        [.. Walk(model.Doc)
            .Where(node => node.Type is "commentStart" or "commentEnd")
            .Select(node => $"{node.Type}:{node.Attrs!["cid"]!.GetValue<string>()}")];

    private static List<string> MarkersOf(string xml) =>
        [.. Regex.Matches(xml, "<w:(commentRangeStart|commentRangeEnd|commentReference) w:id=\"(\\d+)\"")
            .Select(match => $"{match.Groups[1].Value}:{match.Groups[2].Value}")];

    [Fact]
    public void LeAConversaAsRespostasEOResolvido()
    {
        var result = DocxReader.Read(Fixtures.WithCommentThread());
        var comments = result.Model.Comments!;

        Assert.Equal(["0", "1", "2", "3", "4"], comments.Select(comment => comment.Id));
        Assert.Equal("Ana", comments[0].Author);
        Assert.Equal("A", comments[0].Initials);
        Assert.Equal(["Conferir o valor."], comments[0].Paragraphs);
        Assert.StartsWith("2026-03-02T10:00:00", comments[0].Date, StringComparison.Ordinal);
        Assert.Null(comments[0].ParentId);
        Assert.Equal("0", comments[1].ParentId);
        Assert.False(comments[0].Done);
        Assert.True(comments[2].Done);
        Assert.True(comments[4].Rich);
        Assert.False(comments[0].Rich);

        // Uma âncora por conversa: a resposta não vira nó. O de ponto só tem o fim.
        Assert.Equal(
            ["commentStart:0", "commentEnd:0", "commentStart:2", "commentEnd:2", "commentEnd:3", "commentStart:4",
             "commentEnd:4"],
            Anchors(result.Model));

        // Mostrado no painel, o comentário não é aviso nem trava.
        Assert.DoesNotContain(Inventory.Comments, result.Inventory.Invisible);
        Assert.Empty(result.Inventory.Structural);
    }

    [Fact]
    public void EditarOParagrafoComentadoDevolveAsAncorasEAsRespostas()
    {
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));

        Assert.True(EditFirstTextContaining(model, "doze mil", "treze mil"));
        Assert.True(EditFirstTextContaining(model, "Parágrafo com ponto.", "Ponto editado."));

        var (saved, result) = Save(original, model);

        Assert.Equal(2, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);

        // A conversa inteira volta abraçando o mesmo trecho, na ordem do Word; o
        // de ponto volta sem `w:commentRangeEnd`, que ele nunca teve.
        Assert.Equal(
            ["commentRangeStart:0", "commentRangeStart:1", "commentRangeEnd:0", "commentReference:0",
             "commentRangeEnd:1", "commentReference:1",
             "commentRangeStart:2", "commentRangeEnd:2", "commentReference:2",
             "commentReference:3", "commentRangeStart:4", "commentRangeEnd:4", "commentReference:4"],
            MarkersOf(XmlOf(saved)));
        Assert.Contains("treze mil", XmlOf(saved), StringComparison.Ordinal);

        // O corpo dos comentários não mudou: as partes voltam byte a byte.
        var before = PartsOf(original);
        var after = PartsOf(saved);
        Assert.Equal(before["word/comments.xml"], after["word/comments.xml"]);
        Assert.Equal(before["word/commentsExtended.xml"], after["word/commentsExtended.xml"]);

        // E reaberto, o arquivo dá as mesmas âncoras.
        Assert.Equal(Anchors(Open(original)), Anchors(Open(saved)));
    }

    private static bool RemoveFirst(Node node, string type, string cid)
    {
        var content = node.Content ?? [];
        var index = content.FindIndex(child =>
            child.Type == type && child.Attrs?["cid"]?.GetValue<string>() == cid);
        if (index >= 0)
        {
            content.RemoveAt(index);
            return true;
        }

        return content.Any(child => RemoveFirst(child, type, cid));
    }

    [Fact]
    public void ApagarOFimDoTrechoDeixaOComentarioNoPontoDoComeco()
    {
        // O editor apaga o fim com o texto e deixa o começo: um começo sem fim faz
        // o LibreOffice descartar a conversa inteira.
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));
        Assert.True(RemoveFirst(model.Doc, "commentEnd", "0"));

        var (saved, result) = Save(original, model);

        var markers = MarkersOf(XmlOf(saved));
        Assert.DoesNotContain("commentRangeStart:0", markers);
        Assert.DoesNotContain("commentRangeStart:1", markers);
        Assert.Contains("commentReference:0", markers);
        Assert.Contains("commentReference:1", markers);
        Assert.Empty(result.Inventory.Lost);
    }

    [Fact]
    public void ORascunhoGravadoNumPacoteNovoLevaOsComentarios()
    {
        // O rascunho reaberto e gravado num pacote sem comentários: as partes
        // nascem, e a conversa volta inteira — com a resposta ao lado.
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));

        var (saved, result) = Save(Fixtures.Simple(), model);

        Assert.Equal(MarkersOf(XmlOf(original)), MarkersOf(XmlOf(saved)));
        var reopened = DocxReader.Read(saved).Model.Comments!;
        Assert.Equal(["0", "1", "2", "3", "4"], reopened.Select(comment => comment.Id));
        Assert.Equal("0", reopened[1].ParentId);
        Assert.True(reopened[2].Done);
        Assert.Equal(["Conferir o valor."], reopened[0].Paragraphs);
        // O de formatação sai com o texto simples, e a perda é dita.
        Assert.Equal([CommentsWriter.RichEdited], result.Inventory.Lost);
    }

    [Fact]
    public void AbrirESalvarSemEditarNaoReescreveNada()
    {
        var original = Fixtures.WithCommentThread();
        var (saved, result) = Save(original, Clone(Open(original)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Empty(result.Inventory.Lost);

        var before = PartsOf(original);
        var after = PartsOf(saved);
        foreach (var name in before.Keys.Where(name => name != "word/document.xml"))
        {
            Assert.Equal(before[name], after[name]);
        }

        Assert.Equal(MarkersOf(XmlOf(original)), MarkersOf(XmlOf(saved)));
    }

    [Fact]
    public void ORascunhoDeAntesDosComentariosAindaDeclaraAPerda()
    {
        // O rascunho `.sdoc` < 7 não tem as âncoras nos nós. A leitura de
        // referência dele também não as dá — nada muda sem edição —, e reescrever
        // o parágrafo comentado continua perdendo a âncora, como antes.
        var original = Fixtures.WithComment();
        var model = Clone(Open(original)) with { BeforeComments = true };
        foreach (var node in Walk(model.Doc).ToList())
        {
            node.Content?.RemoveAll(child => child.Type is "commentStart" or "commentEnd");
        }

        var (_, untouched) = Save(original, Clone(model));
        Assert.Equal(0, untouched.RewrittenBlocks);

        Assert.True(EditFirstTextContaining(model, "Parágrafo comentado", "Reescrito."));
        var (_, result) = Save(original, model);

        Assert.Contains(result.Inventory.Lost, message => message.Contains("comentário", StringComparison.Ordinal));
    }

    private static Node AnchorNode(string type, string cid) =>
        new() { Type = type, Attrs = new() { ["cid"] = System.Text.Json.Nodes.JsonValue.Create(cid) } };

    /// <summary>O primeiro parágrafo cujo texto contém <paramref name="needle"/>.</summary>
    private static Node ParagraphWith(DocumentModelDto model, string needle) =>
        Walk(model.Doc).First(node => node.Type == "paragraph" &&
                                      (node.Content ?? []).Any(child => child.Text?.Contains(needle, StringComparison.Ordinal) == true));

    private static CommentDto NewComment(string id, string text, string? parentId = null) =>
        new(id, "Vinícius Storch", "2026-10-01T12:00:00Z", [text], false, ParentId: parentId);

    [Fact]
    public void ComentarioNovoNumDocumentoSemComentariosCriaAsPartes()
    {
        var original = Fixtures.Simple();
        var model = Clone(Open(original));
        var paragraph = ParagraphWith(model, "Segundo");
        paragraph.Content!.Insert(0, AnchorNode("commentStart", "0"));
        paragraph.Content.Add(AnchorNode("commentEnd", "0"));
        model = model with { Comments = [NewComment("0", "Revisar.\tCom tab")] };

        var (saved, result) = Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Equal(["commentRangeStart:0", "commentRangeEnd:0", "commentReference:0"], MarkersOf(XmlOf(saved)));
        var parts = PartsOf(saved);
        Assert.Contains("word/comments.xml", parts.Keys);
        Assert.Contains("word/commentsExtended.xml", parts.Keys);
        Assert.Contains("word/people.xml", parts.Keys);
        Assert.Contains("comments+xml", XmlOf(saved, "[Content_Types].xml"), StringComparison.Ordinal);
        Assert.Contains("comments.xml", XmlOf(saved, "word/_rels/document.xml.rels"), StringComparison.Ordinal);
        Assert.Contains("annotationRef", XmlOf(saved, "word/comments.xml"), StringComparison.Ordinal);
        Assert.Contains("w15:author=\"Vinícius Storch\"", XmlOf(saved, "word/people.xml"), StringComparison.Ordinal);

        var reopened = DocxReader.Read(saved);
        var comment = Assert.Single(reopened.Model.Comments!);
        Assert.Equal(["Revisar.\tCom tab"], comment.Paragraphs);
        Assert.Equal("Vinícius Storch", comment.Author);
        Assert.Equal("VS", comment.Initials);
        Assert.NotNull(comment.ParaId);
        Assert.Equal(["commentStart:0", "commentEnd:0"], Anchors(reopened.Model));
    }

    [Fact]
    public void ComentarioDePontoNovoSoTemAReferencia()
    {
        var original = Fixtures.Simple();
        var model = Clone(Open(original));
        ParagraphWith(model, "Terceiro").Content!.Add(AnchorNode("commentEnd", "7"));
        model = model with { Comments = [NewComment("7", "Aqui.")] };

        var (saved, _) = Save(original, model);

        Assert.Equal(["commentReference:7"], MarkersOf(XmlOf(saved)));
        Assert.Equal("7", Assert.Single(DocxReader.Read(saved).Model.Comments!).Id);
    }

    [Fact]
    public void EditarOTextoReescreveSoOComentario()
    {
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));
        model.Comments![0] = model.Comments[0] with { Paragraphs = ["Valor conferido.", "Segunda linha."] };

        var (saved, result) = Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        var before = PartsOf(original);
        var after = PartsOf(saved);
        Assert.NotEqual(before["word/comments.xml"], after["word/comments.xml"]);
        Assert.Equal(before["word/commentsExtended.xml"], after["word/commentsExtended.xml"]);

        var reopened = DocxReader.Read(saved).Model.Comments!;
        Assert.Equal(["Valor conferido.", "Segunda linha."], reopened[0].Paragraphs);
        // O `paraId` é o mesmo: a resposta continua apontando para ele.
        Assert.Equal("10000000", reopened[0].ParaId);
        Assert.Equal("0", reopened[1].ParentId);
        Assert.Equal(["Conferido."], reopened[1].Paragraphs);
    }

    [Fact]
    public void ResolverEReabrirMudaSoOCommentsExtended()
    {
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));
        model.Comments![0] = model.Comments[0] with { Done = true };
        model.Comments[2] = model.Comments[2] with { Done = false };

        var (saved, _) = Save(original, model);

        var before = PartsOf(original);
        var after = PartsOf(saved);
        Assert.Equal(before["word/comments.xml"], after["word/comments.xml"]);
        Assert.NotEqual(before["word/commentsExtended.xml"], after["word/commentsExtended.xml"]);
        var reopened = DocxReader.Read(saved).Model.Comments!;
        Assert.True(reopened[0].Done);
        Assert.False(reopened[2].Done);
    }

    [Fact]
    public void ResponderPoeAsPontasAoLadoDasDoComentario()
    {
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));
        model.Comments!.Add(NewComment("5", "Concordo.", parentId: "0"));

        var (saved, result) = Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(
            ["commentRangeStart:0", "commentRangeStart:1", "commentRangeStart:5", "commentRangeEnd:0",
             "commentReference:0", "commentRangeEnd:1", "commentReference:1", "commentRangeEnd:5",
             "commentReference:5"],
            MarkersOf(XmlOf(saved)).Take(9));

        var reopened = DocxReader.Read(saved).Model.Comments!;
        var reply = Assert.Single(reopened, comment => comment.Id == "5");
        Assert.Equal("0", reply.ParentId);
        Assert.Contains("Vinícius Storch", XmlOf(saved, "word/people.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void ExcluirAConversaLevaAsRespostas()
    {
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));
        Assert.True(RemoveFirst(model.Doc, "commentStart", "0"));
        Assert.True(RemoveFirst(model.Doc, "commentEnd", "0"));
        model = model with { Comments = [.. model.Comments!.Where(comment => comment.Id is not "0" and not "1")] };

        var (saved, result) = Save(original, model);

        Assert.Empty(result.Inventory.Lost);
        var markers = MarkersOf(XmlOf(saved));
        Assert.DoesNotContain(markers, marker => marker.EndsWith(":0", StringComparison.Ordinal) ||
                                                 marker.EndsWith(":1", StringComparison.Ordinal));
        var reopened = DocxReader.Read(saved).Model.Comments!;
        Assert.Equal(["2", "3", "4"], reopened.Select(comment => comment.Id));
        var extended = XmlOf(saved, "word/commentsExtended.xml");
        Assert.DoesNotContain("10000000", extended, StringComparison.Ordinal);
        Assert.DoesNotContain("10000001", extended, StringComparison.Ordinal);
    }

    [Fact]
    public void EditarComentarioComFormatacaoDeclaraAPerda()
    {
        var original = Fixtures.WithCommentThread();
        var model = Clone(Open(original));
        model.Comments![4] = model.Comments[4] with { Paragraphs = ["sem negrito"] };

        var (saved, result) = Save(original, model);

        Assert.Equal([CommentsWriter.RichEdited], result.Inventory.Lost);
        Assert.Equal(["sem negrito"], DocxReader.Read(saved).Model.Comments![4].Paragraphs);
    }
}

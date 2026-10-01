using System.Text.RegularExpressions;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// <summary>
/// Comentários (M10, fase 1): lidos, mostrados e devolvidos ao arquivo.
/// </summary>
/// <remarks>
/// A âncora virou um par de nós no parágrafo, e o corpo do comentário mora fora
/// dos nós. Editar o parágrafo comentado não perde mais nada — e por isso o
/// comentário deixou de travar o documento.
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
    public void PacoteSemComentariosNaoRecebeAncora()
    {
        // O rascunho reaberto e gravado num pacote novo: âncora de comentário que
        // o pacote não tem deixa o arquivo ilegível no LibreOffice.
        var model = Clone(Open(Fixtures.WithCommentThread()));

        var (saved, _) = Save(Fixtures.Simple(), model);

        Assert.Empty(MarkersOf(XmlOf(saved)));
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
}

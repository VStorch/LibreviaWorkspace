using System.Text;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Um comentário como o painel o mostra. Fora dos nós, como os estilos: o corpo
/// mora em <c>word/comments.xml</c>, e o parágrafo leva só as pontas da âncora.
/// </summary>
public sealed record CommentDto(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("author")] string Author,
    [property: JsonPropertyName("date")] string Date,
    [property: JsonPropertyName("paragraphs")] List<string> Paragraphs,
    [property: JsonPropertyName("done")] bool Done,
    // `w15:paraIdParent` traduzido para o id do pai; ausente no que abre a conversa.
    [property: JsonPropertyName("parentId")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? ParentId = null,
    [property: JsonPropertyName("initials")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Initials = null,
    // O `w14:paraId` do último parágrafo, por onde `commentsExtended.xml` liga resposta e "resolvido".
    [property: JsonPropertyName("paraId")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? ParaId = null,
    // Formatação, imagem, campo ou link que o texto simples não mostra: fica no arquivo, e o painel avisa.
    [property: JsonPropertyName("rich")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Rich = false);

/// <summary>
/// <c>word/comments.xml</c> e <c>word/commentsExtended.xml</c> → os comentários do
/// painel. O malformado sai com o que deu para ler.
/// </summary>
public static class CommentsReader
{
    public static List<CommentDto>? Read(MainDocumentPart part)
    {
        var comments = part.WordprocessingCommentsPart?.Comments?.Elements<Comment>().ToList() ?? [];
        if (comments.Count == 0) return null;

        var threads = ThreadsOf(part);
        var result = new List<CommentDto>();
        foreach (var comment in comments)
        {
            if (comment.Id?.Value is not { Length: > 0 } id) continue;

            var paragraphs = comment.Elements<Paragraph>().ToList();
            var paraId = paragraphs.LastOrDefault()?.ParagraphId?.Value;
            var extended = paraId is null ? null : threads.GetValueOrDefault(paraId);

            result.Add(new CommentDto(
                id,
                comment.Author?.Value ?? string.Empty,
                comment.Date?.InnerText ?? string.Empty,
                paragraphs.Select(TextOf).ToList(),
                extended?.Done ?? false,
                ParentId: extended?.Parent,
                Initials: comment.Initials?.Value is { Length: > 0 } initials ? initials : null,
                ParaId: paraId,
                Rich: IsRich(comment)));
        }

        return result.Count == 0 ? null : result;
    }

    /// <summary>
    /// As respostas de cada comentário, na ordem do arquivo. A resposta não vira nó:
    /// a âncora dela é a do pai, e a gravação a devolve ao lado dela.
    /// </summary>
    public static Dictionary<string, List<string>> RepliesOf(MainDocumentPart part)
    {
        var replies = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (var comment in Read(part) ?? [])
        {
            if (comment.ParentId is not { } parent) continue;
            if (!replies.TryGetValue(parent, out var list)) replies[parent] = list = [];
            list.Add(comment.Id);
        }

        return replies;
    }

    internal sealed record ThreadEntry(bool Done, string? Parent);

    /// <summary>`w15:done` e `w15:paraIdParent`, pelo `paraId` do comentário — o pai já traduzido para id.</summary>
    internal static Dictionary<string, ThreadEntry> ThreadsOf(MainDocumentPart part)
    {
        var entries = part.WordprocessingCommentsExPart?.CommentsEx?.Elements<W15.CommentEx>().ToList() ?? [];
        var result = new Dictionary<string, ThreadEntry>(StringComparer.OrdinalIgnoreCase);
        if (entries.Count == 0) return result;

        var idByParaId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var comment in part.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? [])
        {
            if (comment.Id?.Value is { Length: > 0 } id &&
                comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value is { } paraId)
            {
                idByParaId.TryAdd(paraId, id);
            }
        }

        foreach (var entry in entries)
        {
            if (entry.ParaId?.Value is not { } paraId) continue;
            var parent = entry.ParaIdParent?.Value is { } parentParaId
                ? idByParaId.GetValueOrDefault(parentParaId)
                : null;
            result.TryAdd(paraId, new ThreadEntry(entry.Done?.Value == true, parent));
        }

        return result;
    }

    internal static string TextOf(Paragraph paragraph)
    {
        var text = new StringBuilder();
        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case Text piece: text.Append(piece.Text); break;
                case TabChar: text.Append('\t'); break;
                case Break: text.Append('\n'); break;
            }
        }

        return text.ToString();
    }

    private static bool IsRich(Comment comment) =>
        comment.Descendants().Any(element => element is
            DocumentFormat.OpenXml.Wordprocessing.Drawing or Picture or FieldChar or SimpleField or Hyperlink or Table
            or Bold or Italic or Underline or Color or Strike or Highlight or VerticalTextAlignment);
}

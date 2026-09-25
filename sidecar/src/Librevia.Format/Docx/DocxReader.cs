using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O modelo do documento como o editor o entende.
/// </summary>
/// <remarks>
/// Os estilos ficam **fora** dos nós de propósito: a impressão digital de um
/// bloco é feita do que está nele, e um estilo dentro do nó faria todo bloco
/// parecer mudado na gravação seguinte — o documento inteiro reescrito para nada.
///
/// Anuláveis porque o escritor não os usa: um modelo que volta do editor para ser
/// gravado não precisa carregá-los, e `word/styles.xml` volta ao arquivo byte a
/// byte pela gravação cirúrgica. Quem os preenche é a leitura.
/// </remarks>
public sealed record DocumentModelDto(
    [property: JsonPropertyName("page")] PageSetupDto Page,
    [property: JsonPropertyName("doc")] Node Doc,
    [property: JsonPropertyName("styles")] StyleSheetDto? Styles = null,
    // Os blocos vieram achatados — rascunho gravado antes de o leitor passar a
    // levar só a formatação direta. Só a gravação o lê: é o que escolhe a leitura
    // de referência com que os blocos do modelo são comparados.
    [property: JsonPropertyName("flatten")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Flatten = false);

public sealed record OpenResult(
    [property: JsonPropertyName("model")] DocumentModelDto Model,
    [property: JsonPropertyName("inventory")] Inventory Inventory);

/// <summary>
/// Abre um DOCX e produz o que a tela precisa.
/// </summary>
/// <remarks>
/// O pacote **não** fica guardado em lugar nenhum. O sidecar é sem estado por
/// decisão de desenho (docs/02-docx-cirurgico.md): a gravação reabre os bytes
/// originais que o processo main manteve. Guardar o pacote aqui quebraria a
/// promessa da Fase 3.5 de que a morte do sidecar não custa o documento aberto.
/// </remarks>
public static class DocxReader
{
    /// <param name="flatten">
    /// Leva a formatação **efetiva** a todo bloco, como antes de os estilos
    /// virarem CSS. É o que a gravação usa para o rascunho antigo (ver
    /// <see cref="DocxWriter"/>), e o que os testes da cascata usam para ver o
    /// resultado dela num lugar só.
    /// </param>
    public static OpenResult Read(byte[] bytes, bool flatten = false)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = Open(stream);

        var part = document.MainDocumentPart
                   ?? throw new DocxException("O arquivo não contém um documento do Word.");
        var body = part.Document?.Body
                   ?? throw new DocxException("O documento do Word está vazio ou danificado.");

        var inventory = new Inventory();
        NoteWholeDocumentFeatures(part, inventory);

        var (content, _) = new BodyReader(part, inventory, flatten).Read(body);
        var page = PageReader.Read(body, part, inventory);

        var doc = Node.Of("doc");
        doc.Content = content;

        return new OpenResult(new DocumentModelDto(page, doc, StyleReader.Read(part)), inventory);
    }

    /// <summary>
    /// Abre e indexa os blocos — o que a gravação cirúrgica precisa.
    /// </summary>
    public static (WordprocessingDocument Document, MainDocumentPart Part, List<Block> Blocks) Index(
        Stream stream,
        Inventory inventory)
    {
        var document = Open(stream);
        var part = document.MainDocumentPart
                   ?? throw new DocxException("O arquivo não contém um documento do Word.");
        var body = part.Document?.Body
                   ?? throw new DocxException("O documento do Word está vazio ou danificado.");

        var (_, blocks) = new BodyReader(part, inventory).Read(body);
        return (document, part, blocks);
    }

    private static WordprocessingDocument Open(Stream stream)
    {
        try
        {
            return WordprocessingDocument.Open(stream, isEditable: false);
        }
        catch (Exception problem) when (problem is not DocxException)
        {
            // Documento é dado não confiável (spec). Qualquer coisa que a
            // biblioteca lance vira uma frase, nunca um processo derrubado.
            throw new DocxException(
                "Não foi possível abrir este arquivo. Ele pode estar danificado ou não ser um documento do Word.",
                problem);
        }
    }

    /// <summary>
    /// Recursos que existem no pacote inteiro, não num bloco específico.
    /// </summary>
    /// <remarks>
    /// São detectados aqui, e não no corpo, porque vivem em partes separadas.
    /// Todos são <b>invisibilidade</b>: a gravação cirúrgica copia essas partes
    /// intactas, então nada disso se perde — só não aparece na tela.
    /// </remarks>
    private static void NoteWholeDocumentFeatures(MainDocumentPart part, Inventory inventory)
    {
        if (part.WordprocessingCommentsPart?.Comments?.Any() == true)
        {
            inventory.NoteInvisible(Inventory.Comments);
        }

        if (part.FootnotesPart?.Footnotes?.Elements<Footnote>()
                .Any(note => note.Type?.Value is null || note.Type.Value == FootnoteEndnoteValues.Normal) == true)
        {
            inventory.NoteInvisible(Inventory.Footnotes);
        }

        if (part.EndnotesPart?.Endnotes?.Elements<Endnote>()
                .Any(note => note.Type?.Value is null || note.Type.Value == FootnoteEndnoteValues.Normal) == true)
        {
            inventory.NoteInvisible(Inventory.Endnotes);
        }

        var document = part.Document;
        if (document is not null &&
            (document.Descendants<InsertedRun>().Any() || document.Descendants<DeletedRun>().Any()))
        {
            inventory.NoteInvisible(Inventory.TrackedChanges);
        }
    }
}

/// <summary>Falha com frase pronta para o usuário.</summary>
public sealed class DocxException(string message, Exception? inner = null) : Exception(message, inner);

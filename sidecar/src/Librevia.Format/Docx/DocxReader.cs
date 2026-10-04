using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Os estilos ficam **fora** dos nós, pela impressão digital. Anuláveis porque o
/// escritor não os usa: <c>word/styles.xml</c> volta byte a byte.
/// </summary>
public sealed record DocumentModelDto(
    [property: JsonPropertyName("page")] PageSetupDto Page,
    [property: JsonPropertyName("doc")] Node Doc,
    [property: JsonPropertyName("styles")] StyleSheetDto? Styles = null,
    // Rascunho achatado: só a gravação o lê, para escolher a leitura de referência.
    [property: JsonPropertyName("flatten")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Flatten = false,
    // Rascunho anterior às referências (`.sdoc` < 5): só a gravação o lê, como `Flatten`.
    [property: JsonPropertyName("beforeReferences")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeReferences = false,
    // Os que não viraram nó: o editor precisa saber que existem, senão o F9 daria "Erro! Indicador não definido.".
    [property: JsonPropertyName("outsideBookmarks")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<string>? OutsideBookmarks = null,
    // Ver PageReader.ReadAll. Ausente é uma seção só.
    [property: JsonPropertyName("sections")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<PageSetupDto>? Sections = null,
    // Rascunho anterior às seções (`.sdoc` < 6): a página vai só ao `w:sectPr` do corpo.
    [property: JsonPropertyName("beforeSections")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeSections = false,
    // Ver CommentsReader e CommentsWriter. Na gravação, ausente é "não mexa".
    [property: JsonPropertyName("comments")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<CommentDto>? Comments = null,
    // Rascunho anterior aos comentários (`.sdoc` < 7).
    [property: JsonPropertyName("beforeComments")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeComments = false,
    // Só quando ligado; na gravação, ausente é "não mexa" (Revisions.ApplyTracking).
    [property: JsonPropertyName("trackChanges")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    bool? TrackChanges = null,
    // Rascunho anterior às revisões (`.sdoc` < 8).
    [property: JsonPropertyName("beforeRevisions")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeRevisions = false,
    // Ver NotesReader; gravada só quando difere (NotesWriter.ApplyNumbering).
    [property: JsonPropertyName("notes")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    NotesDto? Notes = null,
    // Rascunho anterior às notas (`.sdoc` < 9).
    [property: JsonPropertyName("beforeNotes")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeNotes = false,
    // Rascunho anterior às equações (`.sdoc` < 11).
    [property: JsonPropertyName("beforeMath")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeMath = false,
    // Ver DocumentProperties; cada campo é remendo.
    [property: JsonPropertyName("properties")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    PropertiesDto? Properties = null,
    // Decide o rótulo da parte principal — ver PackageKind.Retarget.
    [property: JsonPropertyName("template")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Template = false);

public sealed record OpenResult(
    [property: JsonPropertyName("model")] DocumentModelDto Model,
    [property: JsonPropertyName("inventory")] Inventory Inventory);

/// <summary>
/// O pacote **não** fica guardado: a gravação reabre os bytes que o main manteve, e
/// a morte do sidecar não custa o documento aberto.
/// </summary>
public static class DocxReader
{
    /// <param name="flatten">A formatação **efetiva** em todo bloco: a leitura do rascunho antigo e dos testes da cascata.</param>
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
        // As macros do `.dotm` não chegam a arquivo nenhum que sair daqui.
        if (PackageKind.HasMacros(bytes)) inventory.NoteLoss(PackageKind.Macros);

        var (content, _) = new BodyReader(part, inventory, flatten).Read(body);
        var (page, sections) = PageReader.ReadAll(body, part, inventory);

        var doc = Node.Of("doc");
        doc.Content = content;

        return new OpenResult(
            new DocumentModelDto(
                page,
                doc,
                StyleReader.Read(part),
                OutsideBookmarks: OutsideBookmarksOf(part, doc),
                Sections: sections,
                Comments: CommentsReader.Read(part),
                TrackChanges: Revisions.TrackingOf(part) ? true : null,
                Notes: NotesReader.Read(part, body),
                Properties: DocumentProperties.Read(document)),
            inventory);
    }

    /// <summary>Os nomes de marcador do pacote que não estão entre os nós lidos.</summary>
    private static List<string>? OutsideBookmarksOf(MainDocumentPart part, Node doc)
    {
        var read = new HashSet<string>(StringComparer.Ordinal);
        void Walk(Node node)
        {
            if (node.Type == "bookmarkStart" && node.Attrs?.GetValueOrDefault("name")?.GetValue<string>() is { } name)
            {
                read.Add(name);
            }

            foreach (var child in node.Content ?? []) Walk(child);
        }

        Walk(doc);

        IEnumerable<DocumentFormat.OpenXml.OpenXmlPartRootElement?> roots =
        [
            part.Document,
            .. part.HeaderParts.Select(header => header.Header),
            .. part.FooterParts.Select(footer => footer.Footer),
            part.FootnotesPart?.Footnotes,
            part.EndnotesPart?.Endnotes,
        ];
        var outside = roots
            .SelectMany(root => root?.Descendants<BookmarkStart>() ?? [])
            .Select(start => start.Name?.Value)
            .OfType<string>()
            .Where(name => name.Length > 0 && !read.Contains(name))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        return outside.Count == 0 ? null : outside;
    }

    /// <summary>O que a gravação cirúrgica precisa.</summary>
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
            // Documento é dado não confiável: o que a biblioteca lance vira frase, e não processo derrubado.
            throw new DocxException(
                "Não foi possível abrir este arquivo. Ele pode estar danificado ou não ser um documento do Word.",
                problem);
        }
    }

    /// <summary>Em partes separadas, e todos invisibilidade: a gravação copia as partes intactas.</summary>
    private static void NoteWholeDocumentFeatures(MainDocumentPart part, Inventory inventory)
    {
        // Sobram as revisões de estrutura, que a gravação de uma tabela ou seção editada
        // perderia, e as de formatação, que se perdem no parágrafo editado.
        var document = part.Document;
        if (document is null) return;

        if (document.Descendants().Any(element =>
                element.LocalName is "cellIns" or "cellDel" or "cellMerge" or "numberingChange"
                    or "sectPrChange" or "tblPrChange"))
        {
            inventory.NoteInvisible(Inventory.StructureRevisions);
        }

        if (document.Descendants<RunPropertiesChange>().Any() ||
            document.Descendants<ParagraphPropertiesChange>().Any())
        {
            inventory.NoteInvisible(Inventory.FormatRevisions);
        }
    }
}

/// <summary>Falha com frase pronta para o usuário.</summary>
public sealed class DocxException(string message, Exception? inner = null) : Exception(message, inner);

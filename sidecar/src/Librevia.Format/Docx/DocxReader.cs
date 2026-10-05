using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Styles stay **outside** the nodes, because of the fingerprint. Nullable because the writer does
/// not use them: <c>word/styles.xml</c> goes back byte for byte.
/// </summary>
public sealed record DocumentModelDto(
    [property: JsonPropertyName("page")] PageSetupDto Page,
    [property: JsonPropertyName("doc")] Node Doc,
    [property: JsonPropertyName("styles")] StyleSheetDto? Styles = null,
    // A flattened draft: only saving reads it, to choose the reference reading.
    [property: JsonPropertyName("flatten")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Flatten = false,
    // A draft older than references (`.sdoc` < 5): only saving reads it, like `Flatten`.
    [property: JsonPropertyName("beforeReferences")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeReferences = false,
    // Those that did not become nodes: the editor must know they exist, or F9 would give "Erro!
    // Indicador não definido.".
    [property: JsonPropertyName("outsideBookmarks")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<string>? OutsideBookmarks = null,
    // See PageReader.ReadAll. Absent means a single section.
    [property: JsonPropertyName("sections")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<PageSetupDto>? Sections = null,
    // A draft older than sections (`.sdoc` < 6): the page only goes to the body's `w:sectPr`.
    [property: JsonPropertyName("beforeSections")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeSections = false,
    // See CommentsReader and CommentsWriter. When saving, absent means "leave alone".
    [property: JsonPropertyName("comments")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<CommentDto>? Comments = null,
    // A draft older than comments (`.sdoc` < 7).
    [property: JsonPropertyName("beforeComments")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeComments = false,
    // Only when on; when saving, absent means "leave alone" (Revisions.ApplyTracking).
    [property: JsonPropertyName("trackChanges")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    bool? TrackChanges = null,
    // A draft older than revisions (`.sdoc` < 8).
    [property: JsonPropertyName("beforeRevisions")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeRevisions = false,
    // See NotesReader; written only when it differs (NotesWriter.ApplyNumbering).
    [property: JsonPropertyName("notes")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    NotesDto? Notes = null,
    // A draft older than notes (`.sdoc` < 9).
    [property: JsonPropertyName("beforeNotes")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeNotes = false,
    // A draft older than equations (`.sdoc` < 11).
    [property: JsonPropertyName("beforeMath")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool BeforeMath = false,
    // See DocumentProperties; each field is a patch.
    [property: JsonPropertyName("properties")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    PropertiesDto? Properties = null,
    // Decides the main part's content type; see PackageKind.Retarget.
    [property: JsonPropertyName("template")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    bool Template = false);

public sealed record OpenResult(
    [property: JsonPropertyName("model")] DocumentModelDto Model,
    [property: JsonPropertyName("inventory")] Inventory Inventory);

/// <summary>
/// The package is **not** kept: saving reopens the bytes main kept, and a sidecar crash does not
/// cost the open document.
/// </summary>
public static class DocxReader
{
    /// <param name="flatten">The **effective** formatting on every block: the reading for old
    /// drafts and for the cascade tests.</param>
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
        // `.dotm` macros reach no file that leaves here.
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

    /// <summary>The package's bookmark names that are not among the read nodes.</summary>
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

    /// <summary>What the surgical save needs.</summary>
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
            // A document is untrusted data: whatever the library throws becomes a sentence, not a
            // crashed process.
            throw new DocxException(
                "Não foi possível abrir este arquivo. Ele pode estar danificado ou não ser um documento do Word.",
                problem);
        }
    }

    /// <summary>In separate parts, and all invisibility: saving copies the parts intact.</summary>
    private static void NoteWholeDocumentFeatures(MainDocumentPart part, Inventory inventory)
    {
        // What remains are structural revisions, which saving an edited table or section would
        // lose, and formatting ones, lost in the edited paragraph.
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

/// <summary>Fails with a sentence ready for the user.</summary>
public sealed class DocxException(string message, Exception? inner = null) : Exception(message, inner);

using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// A document's path in a test: open, change, save, check. <c>Clone</c> through JSON shows the
/// model as the editor returns it, with every attribute materialized.
/// </summary>
internal static class Roundtrip
{
    public static DocumentModelDto Open(byte[] bytes) => DocxReader.Read(bytes).Model;

    /// <summary>
    /// The flattened reading: each block with its effective formatting, style included.
    /// </summary>
    public static DocumentModelDto OpenFlat(byte[] bytes) => DocxReader.Read(bytes, flatten: true).Model;

    /// <summary>
    /// Saves and checks the schema: Word refuses a document outside it, and no assertion on the
    /// model catches that, like the <c>w:right=""</c> LibreOffice tolerates.
    /// </summary>
    public static (byte[] Bytes, SaveResult Result) Save(byte[] original, DocumentModelDto model)
    {
        var saved = DocxWriter.Write(original, model);
        AssertSchema(saved.Bytes);
        return saved;
    }

    /// <summary>The saved document passes the SDK validator without a single error.</summary>
    public static void AssertSchema(byte[] docx)
    {
        using var stream = new MemoryStream(docx.ToArray());
        using var document = WordprocessingDocument.Open(stream, false);

        var errors = new OpenXmlValidator(FileFormatVersions.Office2021)
            .Validate(document)
            .ToList();
        if (errors.Count == 0) return;

        var report = string.Join(
            "\n",
            errors.Select(error => $"[{error.ErrorType}] {error.Path?.XPath}: {error.Description}"));
        Assert.Fail($"O documento gravado está fora do esquema OOXML:\n{report}");
    }

    /// <summary>Clones through JSON, as the model travels.</summary>
    public static DocumentModelDto Clone(DocumentModelDto model) =>
        JsonSerializer.Deserialize<DocumentModelDto>(
            JsonSerializer.Serialize(model, DocxJson.Options), DocxJson.Options)!;

    public static IEnumerable<Node> Walk(Node node)
    {
        yield return node;
        foreach (var child in node.Content ?? [])
        {
            foreach (var deeper in Walk(child)) yield return deeper;
        }
    }

    public static string TextOf(DocumentModelDto model) =>
        string.Concat(Walk(model.Doc).Where(node => node.Type == "text").Select(node => node.Text));

    public static bool EditFirstTextContaining(DocumentModelDto model, string needle, string replacement)
    {
        var target = Walk(model.Doc)
            .FirstOrDefault(node =>
                node.Type == "text" && node.Text?.Contains(needle, StringComparison.Ordinal) == true);
        if (target is null) return false;
        target.Text = replacement;
        return true;
    }

    public static Dictionary<string, byte[]> PartsOf(byte[] docx)
    {
        using var archive = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var stream = entry.Open();
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                return buffer.ToArray();
            },
            StringComparer.Ordinal);
    }

    /// <summary>A part's XML, as text.</summary>
    public static string XmlOf(byte[] docx, string part = "word/document.xml") =>
        Encoding.UTF8.GetString(PartsOf(docx)[part]);
}

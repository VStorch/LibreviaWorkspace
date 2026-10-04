using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// O caminho de um documento num teste: abrir, mexer, gravar, conferir. O
/// <c>Clone</c> pelo JSON mostra o modelo como o editor o devolve, com todo
/// atributo materializado.
/// </summary>
internal static class Roundtrip
{
    public static DocumentModelDto Open(byte[] bytes) => DocxReader.Read(bytes).Model;

    /// <summary>A leitura achatada: cada bloco com a formatação efetiva, estilo incluído.</summary>
    public static DocumentModelDto OpenFlat(byte[] bytes) => DocxReader.Read(bytes, flatten: true).Model;

    /// <summary>
    /// Grava e confere o esquema: o Word recusa o documento fora dele, e nenhuma
    /// assertiva sobre o modelo pega isso, como o <c>w:right=""</c> que o LibreOffice
    /// tolera.
    /// </summary>
    public static (byte[] Bytes, SaveResult Result) Save(byte[] original, DocumentModelDto model)
    {
        var saved = DocxWriter.Write(original, model);
        AssertSchema(saved.Bytes);
        return saved;
    }

    /// <summary>O documento gravado passa pelo validador do SDK sem um erro.</summary>
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

    /// <summary>Clona pelo JSON, como o modelo viaja.</summary>
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

    /// <summary>O XML de uma parte, como texto.</summary>
    public static string XmlOf(byte[] docx, string part = "word/document.xml") =>
        Encoding.UTF8.GetString(PartsOf(docx)[part]);
}

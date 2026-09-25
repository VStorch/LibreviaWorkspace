using System.IO.Compression;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// O caminho que um documento faz num teste: abrir, mexer, gravar, conferir.
/// </summary>
/// <remarks>
/// Mora fora das classes de teste porque são duas a usá-lo — a da promessa
/// central e a dos defeitos de gravação —, e porque o `Clone` pelo JSON não é
/// conveniência: é o único jeito de o teste ver o modelo como o editor o
/// devolve, com todo atributo materializado.
/// </remarks>
internal static class Roundtrip
{
    public static DocumentModelDto Open(byte[] bytes) => DocxReader.Read(bytes).Model;

    /// <summary>A leitura achatada: cada bloco com a formatação efetiva, estilo incluído.</summary>
    public static DocumentModelDto OpenFlat(byte[] bytes) => DocxReader.Read(bytes, flatten: true).Model;

    /// <summary>
    /// Grava — e **confere o esquema** do que foi gravado.
    /// </summary>
    /// <remarks>
    /// A conferência mora aqui, e não em um teste só, porque todo teste de
    /// gravação passa por este método: um documento fora do esquema é um
    /// documento que o Word recusa, e nenhuma assertiva sobre o modelo pega isso.
    ///
    /// Foi o que aconteceu com o recuo. `indentation.Right = right > 0 ? ... :
    /// null` tem tipo `string`, então o `null` virava `StringValue(null)` e o SDK
    /// gravava `w:right=""` — atributo vazio, que o LibreOffice tolera e o Word
    /// não. Os testes comparavam o modelo depois de reler, e reler um atributo
    /// vazio devolve o mesmo modelo: o defeito era invisível para eles e visível
    /// para quem abrisse o arquivo.
    /// </remarks>
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

    /// <summary>Clona pelo JSON — é como o modelo viaja de verdade.</summary>
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

    /// <summary>O XML de uma parte, como texto — é nele que se procura o atributo.</summary>
    public static string XmlOf(byte[] docx, string part = "word/document.xml") =>
        Encoding.UTF8.GetString(PartsOf(docx)[part]);
}

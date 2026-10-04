using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Librevia.Format.Docx;

namespace Librevia.Format.Tests;

/// <summary>
/// Modelos do Word: o <c>.dotx</c> é um <c>.docx</c> com outro rótulo, e a gravação
/// escreve o do destino. Com o rótulo errado, o Word recusa o <c>.docx</c> e abre o
/// <c>.dotx</c> como documento comum.
/// </summary>
public class TemplateTests
{
    private static string ContentTypes(byte[] bytes) => Entry(bytes, "[Content_Types].xml")!;

    /// <summary>O rótulo declarado para `word/document.xml` — a declaração própria vale mais que o padrão.</summary>
    private static string? MainType(byte[] bytes)
    {
        var types = System.Xml.Linq.XDocument.Parse(ContentTypes(bytes).TrimStart('\uFEFF'));
        return types.Root!.Elements()
            .FirstOrDefault(over => over.Name.LocalName == "Override" && (string?)over.Attribute("PartName") == "/word/document.xml")
            ?.Attribute("ContentType")?.Value;
    }

    private static string? Entry(byte[] bytes, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entry = archive.GetEntry(name);
        if (entry is null) return null;
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static byte[] AsTemplate(byte[] docx) => PackageKind.Retarget(docx, template: true, new Inventory());

    /// <summary>
    /// Um `.dotm` mínimo: o pacote do <paramref name="docx"/> com projeto VBA, os
    /// dados dele e o rótulo de modelo com macro.
    /// </summary>
    private static byte[] AsMacroTemplate(byte[] docx)
    {
        using var buffer = new MemoryStream();
        buffer.Write(docx);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            void Put(string name, string content)
            {
                archive.GetEntry(name)?.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }

            string Read(string name)
            {
                using var reader = new StreamReader(archive.GetEntry(name)!.Open());
                return reader.ReadToEnd();
            }

            var types = Read("[Content_Types].xml")
                .Replace(PackageKind.DocumentMain, PackageKind.MacroTemplateMain, StringComparison.Ordinal)
                .Replace(
                    "</Types>",
                    "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>" +
                    "<Override PartName=\"/word/vbaData.xml\" ContentType=\"application/vnd.ms-word.vbaData+xml\"/></Types>",
                    StringComparison.Ordinal);
            Put("[Content_Types].xml", types);

            var rels = Read("word/_rels/document.xml.rels").Replace(
                "</Relationships>",
                "<Relationship Id=\"rIdVba\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/></Relationships>",
                StringComparison.Ordinal);
            Put("word/_rels/document.xml.rels", rels);

            Put("word/vbaProject.bin", "não é VBA de verdade — o conteúdo não importa");
            Put(
                "word/_rels/vbaProject.bin.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/office/2006/relationships/wordVbaData\" Target=\"vbaData.xml\"/></Relationships>");
            Put(
                "word/vbaData.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><wne:vbaSuppData xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\"/>");
        }

        return buffer.ToArray();
    }

    [Fact]
    public void AbrirUmDotxDaOMesmoModeloQueODocxEquivalente()
    {
        var docx = Fixtures.WithHeaderGroup();
        var dotx = AsTemplate(docx);

        Assert.Equal(PackageKind.TemplateMain, MainType(dotx));

        var fromDocx = DocxReader.Read(docx);
        var fromDotx = DocxReader.Read(dotx);
        Assert.Equal(
            JsonSerializer.Serialize(fromDocx, DocxJson.Options),
            JsonSerializer.Serialize(fromDotx, DocxJson.Options));
    }

    [Fact]
    public void GravarODocumentoDeUmModeloEscreveORotuloDeDocumentoSemReescreverNada()
    {
        var dotx = AsTemplate(Fixtures.WithComment());
        var (bytes, result) = Roundtrip.Save(dotx, Roundtrip.Clone(Roundtrip.Open(dotx)));

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.True(result.PreservedBlocks > 0);
        Assert.Empty(result.Inventory.Lost);

        Assert.Equal(PackageKind.DocumentMain, MainType(bytes));
        Assert.DoesNotContain(PackageKind.TemplateMain, ContentTypes(bytes), StringComparison.Ordinal);
        // Fora o rótulo, as partes do modelo saem como estavam.
        Assert.Equal(Entry(dotx, "word/styles.xml"), Entry(bytes, "word/styles.xml"));
        Assert.Equal(Entry(dotx, "word/comments.xml"), Entry(bytes, "word/comments.xml"));
    }

    [Fact]
    public void SalvarComoModeloEscreveORotuloDeModelo()
    {
        var docx = Fixtures.Simple();
        var model = Roundtrip.Clone(Roundtrip.Open(docx)) with { Template = true };
        var (bytes, result) = Roundtrip.Save(docx, model);

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.Equal(PackageKind.TemplateMain, MainType(bytes));

        // E o modelo gravado volta a abrir com o mesmo conteúdo.
        Assert.Equal(Roundtrip.TextOf(Roundtrip.Open(docx)), Roundtrip.TextOf(Roundtrip.Open(bytes)));
    }

    [Fact]
    public void ODocumentoNovoGravadoComoModeloSaiComORotuloDeModelo()
    {
        var created = DocxTemplate.Create(Roundtrip.Open(Fixtures.Simple()).Page);
        var model = Roundtrip.Clone(Roundtrip.Open(Fixtures.Simple())) with { Template = true };
        var (bytes, _) = Roundtrip.Save(created, model);

        Assert.Equal(PackageKind.TemplateMain, MainType(bytes));
    }

    [Fact]
    public void AsMacrosDoDotmSaemEAPerdaEDeclarada()
    {
        var dotm = AsMacroTemplate(Fixtures.WithComment());
        Assert.True(PackageKind.HasMacros(dotm));

        // Declarada já na abertura: quem abre sabe antes de editar.
        var opened = DocxReader.Read(dotm);
        Assert.Contains(PackageKind.Macros, opened.Inventory.Lost);

        var (bytes, result) = Roundtrip.Save(dotm, Roundtrip.Clone(opened.Model));
        Assert.Contains(PackageKind.Macros, result.Inventory.Lost);
        Assert.Equal(0, result.RewrittenBlocks);
        Assert.False(PackageKind.HasMacros(bytes));

        var types = ContentTypes(bytes);
        Assert.Equal(PackageKind.DocumentMain, MainType(bytes));
        Assert.DoesNotContain(PackageKind.MacroTemplateMain, types, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaProject", types, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaData", types, StringComparison.Ordinal);
        Assert.Null(Entry(bytes, "word/vbaProject.bin"));
        Assert.Null(Entry(bytes, "word/vbaData.xml"));
        Assert.Null(Entry(bytes, "word/_rels/vbaProject.bin.rels"));
        Assert.DoesNotContain("vbaProject", Entry(bytes, "word/_rels/document.xml.rels")!, StringComparison.Ordinal);
        Assert.Empty(DocxReader.Read(bytes).Inventory.Lost);
    }

    [Fact]
    public void ODocumentoJaComORotuloCertoVoltaIntacto()
    {
        var docx = Fixtures.Simple();
        Assert.Same(docx, PackageKind.Retarget(docx, template: false, new Inventory()));
    }
}

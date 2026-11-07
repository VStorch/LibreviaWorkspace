using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Librevia.Format.Docx;
using Librevia.Format.Odt;
using Librevia.Format.Protocol;
using static Librevia.Format.Tests.OdtFixtures;

namespace Librevia.Format.Tests;

/// <summary>
/// An ODF template (<c>.ott</c>) differs from a document only in its media type: <c>mimetype</c> and
/// the manifest's root entry. Saving between the two changes those and nothing else.
/// </summary>
public class OdtTemplateTypeTests
{
    private const string TemplateType = "application/vnd.oasis.opendocument.text-template";

    private static readonly byte[] Template = Build("<text:p>Modelo</text:p>", mimeType: TemplateType);
    private static readonly byte[] Document = Build("<text:p>Modelo</text:p>");

    private static DocumentModelDto ModelOf(byte[] odt, bool template) =>
        JsonSerializer.Deserialize<DocumentModelDto>(JsonSerializer.Serialize(OdtReader.Read(odt).Model, JsonOptions.Default), JsonOptions.Default)!
        with { Template = template };

    private static List<(string Name, byte[] Data, bool Stored)> Entries(byte[] zip)
    {
        using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        return [.. archive.Entries.Select(entry =>
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return (entry.FullName, copy.ToArray(), entry.CompressedLength == entry.Length);
        })];
    }

    private static string Part(byte[] zip, string name) => Encoding.UTF8.GetString(Entries(zip).Single(entry => entry.Name == name).Data);

    private static void AssertType(byte[] zip, string type)
    {
        var first = Entries(zip)[0];
        Assert.Equal("mimetype", first.Name);
        Assert.True(first.Stored);
        Assert.Equal(type, Encoding.ASCII.GetString(first.Data));
        Assert.Contains($"manifest:full-path=\"/\" manifest:media-type=\"{type}\"", Part(zip, "META-INF/manifest.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void ATemplateOpensAsAnyDocument()
    {
        var text = OdtReader.Read(Template).Model.Doc.Content!.Single().Content!.Single().Text;

        Assert.Equal("Modelo", text);
    }

    [Fact]
    public void SavingAsTemplateWritesTheTemplateTypeAndKeepsTheOtherParts()
    {
        var (saved, _) = OdtWriter.Write(Document, ModelOf(Document, template: true));

        AssertType(saved, TemplateType);
        Assert.Equal(Part(Document, "content.xml"), Part(saved, "content.xml"));
        Assert.Equal(Part(Document, "styles.xml"), Part(saved, "styles.xml"));
    }

    [Fact]
    public void SavingATemplateAsADocumentGoesBackToText()
    {
        var (saved, _) = OdtWriter.Write(Template, ModelOf(Template, template: false));

        AssertType(saved, MimeType);
    }

    [Fact]
    public void KeepingTheTypeLeavesTheManifestBytes()
    {
        var (saved, _) = OdtWriter.Write(Template, ModelOf(Template, template: true));

        AssertType(saved, TemplateType);
        Assert.Equal(Part(Template, "META-INF/manifest.xml"), Part(saved, "META-INF/manifest.xml"));
    }

    [Fact]
    public void CreatingATemplateWritesTheTemplateType()
    {
        var page = ModelOf(Document, template: false).Page;

        AssertType(OdtTemplate.Create(new OdtCreateDto(page, Template: true)), TemplateType);
        AssertType(OdtTemplate.Create(new OdtCreateDto(page)), MimeType);
    }
}

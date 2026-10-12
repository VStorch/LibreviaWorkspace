using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Librevia.Format.Docx;
using Librevia.Format.Privacy;
using Librevia.Format.Xlsx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// The personal data inspector: it finds names, folders and image metadata, and removes them
/// touching only the parts that carry them.
public class PrivacyTests
{
    private const string Anonymous = "Autor";
    private const string TemplatePath = @"C:\Users\ana.lima\Modelos\Ofício.dotx";
    private const string ExcelFolder = @"C:\Users\ana.lima\Documents\";

    private static readonly PrivacyRequest Inspection = new(Anonymous);

    private static PrivacyRequest Remove(params string[] categories) => new(Anonymous, [.. categories]);

    /// <summary>
    /// A photo as a phone saves it: EXIF with the orientation and the GPS position, a comment, and
    /// the pixels after the scan start.
    /// </summary>
    internal static byte[] PhotoWithExif(ushort orientation)
    {
        var tiff = new byte[8 + 2 + (2 * 12) + 4];
        "II"u8.CopyTo(tiff);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(2), 42);
        BinaryPrimitives.WriteUInt32LittleEndian(tiff.AsSpan(4), 8);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(8), 2);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(10), 0x0112);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(12), 3);
        BinaryPrimitives.WriteUInt32LittleEndian(tiff.AsSpan(14), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(18), orientation);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(22), 0x8825);
        BinaryPrimitives.WriteUInt16LittleEndian(tiff.AsSpan(24), 4);
        BinaryPrimitives.WriteUInt32LittleEndian(tiff.AsSpan(26), 1);

        using var image = new MemoryStream();
        image.Write([0xFF, 0xD8]);
        Segment(image, 0xE0, "JFIF\0\x01\x01\0\0\x01\0\x01\0\0"u8.ToArray());
        Segment(image, 0xE1, [.. "Exif\0\0"u8, .. tiff]);
        Segment(image, 0xFE, "GPS -23.55, -46.63"u8.ToArray());
        Segment(image, 0xDB, new byte[65]);
        image.Write([0xFF, 0xDA, 0x00, 0x02, 0x12, 0x34, 0x56, 0xFF, 0xD9]);
        return image.ToArray();
    }

    private static void Segment(Stream image, byte marker, byte[] payload)
    {
        image.Write([0xFF, marker, (byte)((payload.Length + 2) >> 8), (byte)(payload.Length + 2)]);
        image.Write(payload);
    }

    /// <summary>A Word document with names everywhere Word puts them.</summary>
    private static byte[] DocumentWithPersonalData()
    {
        using var buffer = new MemoryStream();
        buffer.Write(Fixtures.WithCommentThread());
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            PropertiesTests.Write(document.AddCoreFilePropertiesPart(), PropertiesTests.Core);
            PropertiesTests.Write(document.AddExtendedFilePropertiesPart(), PropertiesTests.App);
            PropertiesTests.Write(document.AddCustomFilePropertiesPart(), PropertiesTests.Custom);

            var thumbnail = document.AddThumbnailPart("image/jpeg");
            using (var stream = thumbnail.GetStream()) stream.Write(PhotoWithExif(1));

            var main = document.MainDocumentPart!;
            var settings = main.AddNewPart<DocumentSettingsPart>();
            var template = settings.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
                new Uri("file:///" + TemplatePath.Replace('\\', '/'), UriKind.Absolute));
            settings.Settings = new Settings(new Zoom { Percent = "100" }, new AttachedTemplate { Id = template.Id });

            var people = main.AddNewPart<WordprocessingPeoplePart>();
            PropertiesTests.Write(people,
                """<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">""" +
                """<w15:person w15:author="Ana"><w15:presenceInfo w15:providerId="AD" w15:userId="S-1-5-21-ana"/></w15:person></w15:people>""");

            var photo = main.AddImagePart(ImagePartType.Jpeg);
            using (var stream = photo.GetStream()) stream.Write(PhotoWithExif(6));
        }

        return buffer.ToArray();
    }

    private static byte[] PartBytes(byte[] package, string part)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(part)!.Open();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private static string? ImagePartName(byte[] docx) =>
        PartsOf(docx).Keys.FirstOrDefault(name =>
            name.StartsWith("media/", StringComparison.Ordinal) && Path.GetExtension(name) is ".jpg" or ".png");

    [Fact]
    public void InspecaoEncontraNomesDasPropriedadesDosComentariosEDasPessoas()
    {
        var report = PersonalData.Inspect(DocumentWithPersonalData(), Inspection);

        Assert.Equal(["Ana Lima", "Bia", "Carla", "ACME", "Ana", "Bruno", "Dora"], report.Names);
        Assert.Equal([TemplatePath.Replace('\\', '/')], report.Paths.Select(path => path.Replace("file:///", "")));
        Assert.Equal(["Cliente"], report.CustomProperties);
        Assert.True(report.Thumbnail);
        Assert.Equal(1, report.ImagesWithMetadata);
        Assert.False(report.RemovePersonalInformation);
    }

    [Fact]
    public void RemoverNomesTrocaOsAutoresEApagaAAutoria()
    {
        var clean = PersonalData.Scrub(DocumentWithPersonalData(), Remove(PrivacyCategory.Names)).Bytes;

        var model = Open(clean);
        Assert.All(model.Comments!, comment => Assert.Equal(Anonymous, comment.Author));
        Assert.Null(model.Properties!.Creator);
        Assert.Null(model.Properties.LastModifiedBy);
        Assert.Null(model.Properties.Company);
        Assert.Null(model.Properties.Manager);
        Assert.Equal("Relatório anual", model.Properties.Title);
        Assert.True(model.Properties.RemovePersonalInformation);

        var people = XmlOf(clean, "word/people.xml");
        Assert.DoesNotContain("presenceInfo", people);
        Assert.Contains("w15:author=\"Autor\"", people);
        Assert.Contains("w:initials=\"A\"", XmlOf(clean, "word/comments.xml"));
        Assert.Empty(PersonalData.Inspect(clean, Inspection).Names);
        AssertSchema(clean);
    }

    [Fact]
    public void RemoverNomesTrocaOAutorDasAlteracoesControladas()
    {
        var clean = PersonalData.Scrub(Fixtures.WithTrackedChanges(), Remove(PrivacyCategory.Names)).Bytes;

        Assert.DoesNotContain("Revisor", XmlOf(clean));
        Assert.Contains("w:author=\"Autor\"", XmlOf(clean));
        Assert.Contains("<w:removePersonalInformation", XmlOf(clean, "word/settings.xml"));
    }

    [Fact]
    public void PartesSemDadosPessoaisVoltamComOsMesmosBytes()
    {
        var original = DocumentWithPersonalData();
        var clean = PersonalData.Scrub(original, Remove(PrivacyCategory.Thumbnail)).Bytes;

        var before = PartsOf(original);
        var after = PartsOf(clean);
        Assert.Equal(before.Count - 1, after.Count);
        foreach (var (name, bytes) in after)
        {
            if (name is "_rels/.rels" or "[Content_Types].xml") continue;
            Assert.Equal(before[name], bytes);
        }
    }

    [Fact]
    public void RemoveMiniaturaEPropriedadesPersonalizadasComRelacaoETipo()
    {
        var clean = PersonalData.Scrub(
            DocumentWithPersonalData(),
            Remove(PrivacyCategory.Thumbnail, PrivacyCategory.CustomProperties)).Bytes;

        var parts = PartsOf(clean);
        Assert.DoesNotContain("docProps/custom.xml", parts.Keys);
        Assert.DoesNotContain(parts.Keys, name => name.StartsWith("docProps/thumbnail", StringComparison.Ordinal));
        Assert.DoesNotContain("custom.xml", XmlOf(clean, "_rels/.rels"));
        Assert.DoesNotContain("custom.xml", XmlOf(clean, "[Content_Types].xml"));
        using var reopened = WordprocessingDocument.Open(new MemoryStream(clean), false);
        Assert.Null(reopened.CustomFilePropertiesPart);
        Assert.Null(reopened.ThumbnailPart);
    }

    [Fact]
    public void RemoveOCaminhoDoModeloAnexado()
    {
        var clean = PersonalData.Scrub(DocumentWithPersonalData(), Remove(PrivacyCategory.Paths)).Bytes;

        Assert.DoesNotContain("attachedTemplate", XmlOf(clean, "word/settings.xml"));
        Assert.DoesNotContain("ana.lima", XmlOf(clean, "word/_rels/settings.xml.rels"));
        Assert.Contains("w:zoom", XmlOf(clean, "word/settings.xml"));
        Assert.Empty(PersonalData.Inspect(clean, Inspection).Paths);
    }

    [Fact]
    public void ImagemPerdeOExifEOComentarioMasGuardaAOrientacao()
    {
        var original = DocumentWithPersonalData();
        var clean = PersonalData.Scrub(original, Remove(PrivacyCategory.ImageMetadata)).Bytes;

        var name = ImagePartName(clean)!;
        var photo = PartBytes(clean, name);
        Assert.True(photo.Length < PartBytes(original, name).Length);
        Assert.Equal(-1, photo.AsSpan().IndexOf("GPS"u8));
        Assert.Equal(-1, photo.AsSpan().IndexOf("II*"u8));
        Assert.True(photo.AsSpan().IndexOf("JFIF"u8) > 0);
        Assert.True(photo.AsSpan().EndsWith((byte[])[0x12, 0x34, 0x56, 0xFF, 0xD9]));

        var exif = photo.AsSpan().IndexOf("Exif\0\0MM"u8);
        Assert.True(exif > 0);
        Assert.Equal(6, BinaryPrimitives.ReadUInt16BigEndian(photo.AsSpan(exif + 6 + 18)));
        Assert.Equal(0, PersonalData.Inspect(clean, Inspection).ImagesWithMetadata);
    }

    [Fact]
    public void FotoEmPeNaoGanhaExifNovo()
    {
        using var buffer = new MemoryStream();
        buffer.Write(Fixtures.Simple());
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            var photo = document.MainDocumentPart!.AddImagePart(ImagePartType.Jpeg);
            using var stream = photo.GetStream();
            stream.Write(PhotoWithExif(1));
        }

        var clean = PersonalData.Scrub(buffer.ToArray(), Remove(PrivacyCategory.ImageMetadata)).Bytes;

        Assert.Equal(-1, PartBytes(clean, ImagePartName(clean)!).AsSpan().IndexOf("Exif"u8));
    }

    [Fact]
    public void PngPerdeOsTextosEContinuaValido()
    {
        static byte[] Chunk(string type, byte[] data)
        {
            var chunk = new byte[12 + data.Length];
            BinaryPrimitives.WriteInt32BigEndian(chunk, data.Length);
            Encoding.ASCII.GetBytes(type).CopyTo(chunk, 4);
            data.CopyTo(chunk, 8);
            return chunk;
        }

        byte[] png =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            .. Chunk("IHDR", new byte[13]),
            .. Chunk("tEXt", "Author\0Ana Lima"u8.ToArray()),
            .. Chunk("IDAT", [1, 2, 3]),
            .. Chunk("IEND", []),
        ];
        using var buffer = new MemoryStream();
        buffer.Write(Fixtures.Simple());
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            var image = document.MainDocumentPart!.AddImagePart(ImagePartType.Png);
            using var stream = image.GetStream();
            stream.Write(png);
        }

        var clean = PersonalData.Scrub(buffer.ToArray(), Remove(PrivacyCategory.ImageMetadata)).Bytes;
        var stripped = PartBytes(clean, ImagePartName(clean)!);

        Assert.Equal(png.Length - 12 - 15, stripped.Length);
        Assert.Equal(-1, stripped.AsSpan().IndexOf("Ana Lima"u8));
        Assert.True(stripped.AsSpan().IndexOf("IDAT"u8) > 0);
    }

    [Fact]
    public void SemCategoriasNadaMuda()
    {
        var original = DocumentWithPersonalData();

        var (bytes, changes) = PersonalData.Scrub(original, Remove());

        Assert.Same(original, bytes);
        Assert.Equal(0, changes);
    }

    [Fact]
    public void GravarComAMarcaNoModeloAEscreveNoSettings()
    {
        var original = Fixtures.Simple();
        var model = Open(original) with { Properties = new PropertiesDto(RemovePersonalInformation: true) };

        var saved = Save(original, model).Bytes;

        Assert.Contains("<w:removePersonalInformation", XmlOf(saved, "word/settings.xml"));
        Assert.True(Open(saved).Properties!.RemovePersonalInformation);
    }

    /// <summary>A workbook with a note, the folder Excel saved it in and the author.</summary>
    private static byte[] WorkbookWithPersonalData()
    {
        using var book = new XLWorkbook();
        book.Author = "Ana Lima";
        var sheet = book.Worksheets.Add("Plan1");
        sheet.Cell("A1").Value = 1;
        sheet.Cell("A1").CreateComment().SetAuthor("Ana Lima").AddSignature().AddText("conferir");

        using var stream = new MemoryStream();
        book.SaveAs(stream);

        const string folder =
            """<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">""" +
            """<mc:Choice Requires="x15"><x15ac:absPath xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" url=""" +
            "\"" + ExcelFolder + "\"/></mc:Choice></mc:AlternateContent>";
        return RewritePart(stream.ToArray(), "xl/workbook.xml", xml => xml.Replace("<x:sheets>", folder + "<x:sheets>")
            .Replace("<sheets>", folder + "<sheets>"));
    }

    private static byte[] RewritePart(byte[] package, string part, Func<string, string> change)
    {
        using var buffer = new MemoryStream();
        buffer.Write(package);
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(part)!;
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(part).Open());
            writer.Write(change(xml));
        }

        return buffer.ToArray();
    }

    [Fact]
    public void PlanilhaPerdeOAutorDaNotaEAPasta()
    {
        var original = WorkbookWithPersonalData();
        var report = PersonalData.Inspect(original, Inspection);
        Assert.Contains("Ana Lima", report.Names);
        Assert.Equal([ExcelFolder], report.Paths);

        var clean = PersonalData.Scrub(original, Remove(PrivacyCategory.Names, PrivacyCategory.Paths)).Bytes;

        var comments = XlsxFixtures.PartText(clean, "xl/comments1.xml")!;
        Assert.DoesNotContain("Ana Lima", comments);
        Assert.Contains("Autor:", comments);
        var workbook = XlsxFixtures.PartText(clean, "xl/workbook.xml")!;
        Assert.DoesNotContain("absPath", workbook);
        Assert.Contains("filterPrivacy=\"1\"", workbook);
        Assert.True(XlsxReader.Read(clean).Workbook.Properties!.RemovePersonalInformation);
        Assert.Empty(PersonalData.Inspect(clean, Inspection).Names);
    }

    [Fact]
    public void PlanilhaGravadaGuardaAMarca()
    {
        var clean = PersonalData.Scrub(WorkbookWithPersonalData(), Remove(PrivacyCategory.Names)).Bytes;
        var model = XlsxReader.Read(clean).Workbook;

        var saved = XlsxWriter.Write(clean, model).Bytes;

        Assert.True(XlsxReader.Read(saved).Workbook.Properties!.RemovePersonalInformation);
    }

    [Fact]
    public void ArquivoQueNaoEPacoteERecusado()
    {
        Assert.Throws<PrivacyException>(() => PersonalData.Inspect("não é zip"u8.ToArray(), Inspection));
    }
}

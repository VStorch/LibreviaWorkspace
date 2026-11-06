using System.IO.Compression;
using System.Text;
using Librevia.Format.Odt;
using Librevia.Format.Privacy;
using static Librevia.Format.Tests.OdtFixtures;

namespace Librevia.Format.Tests;

/// <summary>
/// The personal data inspector on an ODT: the same categories as for DOCX, found in
/// <c>meta.xml</c>, in comments and tracked changes, in the thumbnail and in the pictures; cleaning
/// leaves every other entry with the same bytes.
/// </summary>
public class OdtPrivacyTests
{
    private const string Anonymous = "Autor";
    private const string TemplateHref = "file:///home/ana.lima/Modelos/Of%C3%ADcio.ott";

    private static readonly PrivacyRequest Inspection = new(Anonymous);

    private static PrivacyRequest Remove(params string[] categories) => new(Anonymous, [.. categories]);

    private const string Body =
        """<text:tracked-changes><text:changed-region text:id="ct1"><text:insertion><office:change-info><dc:creator>Bruno Reis</dc:creator><dc:date>2026-01-01T00:00:00</dc:date></office:change-info></text:insertion></text:changed-region></text:tracked-changes>""" +
        """<text:p>Texto<office:annotation><dc:creator>Carla Dias</dc:creator><meta:creator-initials>CD</meta:creator-initials><text:p>Nota</text:p></office:annotation></text:p>""";

    private const string Meta =
        """<meta:initial-creator>Ana Lima</meta:initial-creator><dc:creator>Ana Lima</dc:creator><meta:printed-by>Ana L.</meta:printed-by>""" +
        $"""<meta:template xlink:type="simple" xlink:href="{TemplateHref}" xlink:title="Ofício"/>""" +
        """<meta:user-defined meta:name="Cliente">ACME</meta:user-defined><meta:user-defined meta:name="Company">Prefeitura</meta:user-defined>""";

    private static byte[] Odt(string body = Body, string? meta = Meta) => Build(body, meta: meta, files: new Dictionary<string, byte[]>
    {
        ["Pictures/foto.jpg"] = PrivacyTests.PhotoWithExif(1),
        ["Thumbnails/thumbnail.png"] = Png(),
    });

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

    private static byte[] Bytes(byte[] zip, string name) => Entries(zip).Single(entry => entry.Name == name).Data;

    private static string Part(byte[] zip, string name) => Encoding.UTF8.GetString(Bytes(zip, name));

    [Fact]
    public void InspectionFindsWhatAnOdtRevealsAboutItsAuthors()
    {
        var report = PersonalData.Inspect(Odt(), Inspection);

        Assert.Equal(["Ana Lima", "Ana L.", "Prefeitura", "Bruno Reis", "Carla Dias"], report.Names);
        Assert.Equal(["file:///home/ana.lima/Modelos/Ofício.ott"], report.Paths);
        Assert.Equal(["Cliente"], report.CustomProperties);
        Assert.True(report.Thumbnail);
        Assert.Equal(1, report.ImagesWithMetadata);
        Assert.False(report.RemovePersonalInformation);
    }

    [Fact]
    public void RemovingNamesDropsTheMetaAuthorsSignsTheRestAnonymouslyAndLeavesTheMark()
    {
        var original = Odt();
        var (clean, changes) = PersonalData.Scrub(original, Remove(PrivacyCategory.Names));

        Assert.True(changes > 0);
        var report = PersonalData.Inspect(clean, Inspection);
        Assert.Empty(report.Names);
        Assert.True(report.RemovePersonalInformation);
        Assert.Contains("<dc:creator>Autor</dc:creator>", Part(clean, "content.xml"), StringComparison.Ordinal);
        Assert.DoesNotContain("creator", Part(clean, "meta.xml"), StringComparison.Ordinal);
        Assert.DoesNotContain("printed-by", Part(clean, "meta.xml"), StringComparison.Ordinal);
        Assert.DoesNotContain("Prefeitura", Part(clean, "meta.xml"), StringComparison.Ordinal);
        Assert.Contains("ACME", Part(clean, "meta.xml"), StringComparison.Ordinal);
        Assert.Contains("<meta:creator-initials>A</meta:creator-initials>", Part(clean, "content.xml"), StringComparison.Ordinal);
        Assert.Contains("full-path=\"settings.xml\"", Part(clean, "META-INF/manifest.xml"), StringComparison.Ordinal);
        Assert.Equal(Bytes(original, "styles.xml"), Bytes(clean, "styles.xml"));
        Assert.True(OdtReader.Read(clean).Model.Properties!.RemovePersonalInformation);
    }

    [Fact]
    public void TheMarkGoesIntoAnExistingSettingsPart()
    {
        const string settings =
            """<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item config:name="ViewAreaTop" config:type="long">0</config:config-item></config:config-item-set></office:settings></office:document-settings>""";
        var original = Build("<text:p>Texto</text:p>", meta: Meta, files: new Dictionary<string, byte[]> { ["settings.xml"] = Utf8(settings) });

        var (clean, _) = PersonalData.Scrub(original, Remove(PrivacyCategory.Names));

        var written = Part(clean, "settings.xml");
        Assert.Contains("ViewAreaTop", written, StringComparison.Ordinal);
        Assert.Contains("LibreviaRemovePersonalInformation", written, StringComparison.Ordinal);
        Assert.Equal(Bytes(original, "META-INF/manifest.xml"), Bytes(clean, "META-INF/manifest.xml"));
    }

    [Fact]
    public void CleaningKeepsTheOtherEntriesByteForByteAndTheMimeTypeFirstAndStored()
    {
        var original = Odt();
        var (clean, changes) = PersonalData.Scrub(original, Remove(PrivacyCategory.ImageMetadata));

        Assert.Equal(1, changes);
        var first = Entries(clean)[0];
        Assert.Equal(("mimetype", true), (first.Name, first.Stored));
        Assert.NotEqual(Bytes(original, "Pictures/foto.jpg"), Bytes(clean, "Pictures/foto.jpg"));
        foreach (var (name, data, _) in Entries(original).Where(entry => entry.Name != "Pictures/foto.jpg"))
        {
            Assert.Equal(data, Bytes(clean, name));
        }

        Assert.Equal(Entries(original).Select(entry => entry.Name), Entries(clean).Select(entry => entry.Name));
    }

    [Fact]
    public void RemovingPathsCustomPropertiesAndTheThumbnailLeavesNothingToReportAndTheCompanyToNames()
    {
        var (clean, _) = PersonalData.Scrub(Odt(), Remove(
            PrivacyCategory.Paths, PrivacyCategory.CustomProperties, PrivacyCategory.Thumbnail));

        var report = PersonalData.Inspect(clean, Inspection);
        Assert.Empty(report.Paths);
        Assert.Empty(report.CustomProperties);
        Assert.False(report.Thumbnail);
        Assert.DoesNotContain("Thumbnails/", Part(clean, "META-INF/manifest.xml"), StringComparison.Ordinal);
        Assert.Contains("Prefeitura", Part(clean, "meta.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void AnOdtWithNothingPersonalComesBackUntouched()
    {
        var original = Build("<text:p>Texto</text:p>");

        var (clean, changes) = PersonalData.Scrub(original, Remove(
            PrivacyCategory.Paths, PrivacyCategory.CustomProperties, PrivacyCategory.Thumbnail, PrivacyCategory.ImageMetadata));

        Assert.Equal(0, changes);
        Assert.Same(original, clean);
    }

    [Fact]
    public void AZipWithoutTheMimeTypeIsAnOdtOnlyWithTheOdfParts()
    {
        var notes = Zip([("notas.txt", Utf8("Ana Lima"))]);
        var untyped = Build("<text:p>Texto</text:p>", meta: Meta, mimeType: null);

        Assert.IsNotType<OdtException>(Record.Exception(() => PersonalData.Inspect(notes, Inspection)));
        Assert.Contains("Ana Lima", PersonalData.Inspect(untyped, Inspection).Names);
    }
}

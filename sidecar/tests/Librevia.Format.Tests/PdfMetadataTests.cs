using System.Text;
using Librevia.Format.Docx;
using Librevia.Format.Pdf;

namespace Librevia.Format.Tests;

/// The information dictionary and XMP, changed by an incremental update that leaves the original
/// bytes as they were.
public class PdfMetadataTests
{
    private static byte[] Saved(byte[] original, PropertiesDto patch) => [.. original, .. PdfMetadata.Update(original, patch)];

    private static string Text(byte[] bytes) => Encoding.Latin1.GetString(bytes);

    [Fact]
    public void ReadsThePropertiesOfTheInformationDictionary()
    {
        var inspection = PdfMetadata.Inspect(PdfFixtures.Classic());

        Assert.Equal("1.4", inspection.Version);
        Assert.False(inspection.Encrypted);
        Assert.False(inspection.Signed);
        Assert.Equal("Relatório anual", inspection.Properties?.Title);
        Assert.Equal("Ana Lima", inspection.Properties?.Creator);
        Assert.Equal("Contas", inspection.Properties?.Subject);
        Assert.Equal("2025-01-02T03:04:05-03:00", inspection.Properties?.Created);
        Assert.Equal("2025-02-03T04:05:06Z", inspection.Properties?.Modified);
        Assert.Equal("Editor", inspection.Application);
        Assert.Equal("Gerador 1.0", inspection.Producer);
    }

    [Fact]
    public void WithoutDictionaryReadsTheXmp()
    {
        var inspection = PdfMetadata.Inspect(PdfFixtures.Classic(info: null, xmp: PdfFixtures.Xmp));

        Assert.Equal("Relatório anual", inspection.Properties?.Title);
        Assert.Equal("Ana Lima", inspection.Properties?.Creator);
        Assert.Equal("Gerador 1.0", inspection.Producer);
    }

    [Fact]
    public void UpdateComesAfterTheIntactOriginal()
    {
        var original = PdfFixtures.Classic();

        var update = PdfMetadata.Update(original, new PropertiesDto(Title: "Relatório revisto"));
        var saved = Saved(original, new PropertiesDto(Title: "Relatório revisto"));

        Assert.NotEmpty(update);
        Assert.Equal(original, saved[..original.Length]);
        Assert.Contains("/Prev ", Text(update));
        Assert.EndsWith("%%EOF\n", Text(update));
    }

    [Fact]
    public void ReopeningShowsTheNewFieldsAndKeepsTheOthers()
    {
        var saved = Saved(PdfFixtures.Classic(), new PropertiesDto(Title: "Relatório revisto", Keywords: "contas; 2026", Subject: ""));

        var inspection = PdfMetadata.Inspect(saved);
        Assert.Equal("Relatório revisto", inspection.Properties?.Title);
        Assert.Equal("contas; 2026", inspection.Properties?.Keywords);
        Assert.Null(inspection.Properties?.Subject);
        Assert.Equal("Ana Lima", inspection.Properties?.Creator);
        Assert.Equal("Editor", inspection.Application);
        Assert.Contains("/Custom", Text(saved[PdfFixtures.Classic().Length..]));
    }

    [Fact]
    public void ModifiedDateGoesInPdfFormat()
    {
        var saved = Saved(PdfFixtures.Classic(), new PropertiesDto(Modified: "2026-10-12T16:00:00Z"));

        // Strings go in hex, which needs no escaping.
        Assert.Contains(Convert.ToHexString(Encoding.ASCII.GetBytes("D:20261012160000Z")), Text(saved));
        Assert.Equal("2026-10-12T16:00:00Z", PdfMetadata.Inspect(saved).Properties?.Modified);
    }

    [Fact]
    public void WithoutChangeAddsNothing()
    {
        var update = PdfMetadata.Update(PdfFixtures.Classic(), new PropertiesDto(Title: "Relatório anual", Creator: "Ana Lima"));

        Assert.Empty(update);
    }

    [Fact]
    public void PdfWithoutDictionaryGetsANewOne()
    {
        var original = PdfFixtures.Classic(info: null);

        var saved = Saved(original, new PropertiesDto(Title: "Novo"));

        Assert.Equal("Novo", PdfMetadata.Inspect(saved).Properties?.Title);
        Assert.Contains("/Info 5 0 R", Text(saved[original.Length..]));
    }

    [Fact]
    public void XmpChangesTogetherAndKeepsWhatItDoesNotKnow()
    {
        var original = PdfFixtures.Classic(xmp: PdfFixtures.Xmp);

        var saved = Saved(original, new PropertiesDto(Title: "Relatório revisto", Creator: "Bia", Modified: "2026-10-12T16:00:00Z"));

        var update = Encoding.UTF8.GetString(saved[original.Length..]);
        Assert.Contains("""<rdf:li xml:lang="x-default">Relatório revisto</rdf:li>""", update);
        Assert.Contains("<rdf:li>Bia</rdf:li>", update);
        Assert.Contains("""xmp:ModifyDate="2026-10-12T16:00:00Z""", update);
        Assert.Contains("<xmp:MetadataDate>2026-10-12T16:00:00Z</xmp:MetadataDate>", update);
        Assert.Contains("""pdfaid:part="2""", update);
        Assert.Contains("<?xpacket end=\"w\"?>", update);
        Assert.Contains("5 0 obj", update);
    }

    [Fact]
    public void ReadsAndUpdatesAPdfWithXrefStream()
    {
        var original = PdfFixtures.WithXrefStream();
        Assert.Equal("Relatório anual", PdfMetadata.Inspect(original).Properties?.Title);

        var saved = Saved(original, new PropertiesDto(Title: "Relatório revisto"));

        var update = Text(saved[original.Length..]);
        Assert.Contains("/Type /XRef", update);
        Assert.DoesNotContain("trailer", update);
        Assert.Equal("Relatório revisto", PdfMetadata.Inspect(saved).Properties?.Title);
        Assert.Equal("Contas", PdfMetadata.Inspect(saved).Properties?.Subject);
    }

    [Fact]
    public void TwoConsecutiveUpdatesChain()
    {
        var once = Saved(PdfFixtures.Classic(), new PropertiesDto(Title: "Primeira"));
        var twice = Saved(once, new PropertiesDto(Keywords: "segunda"));

        var inspection = PdfMetadata.Inspect(twice);
        Assert.Equal("Primeira", inspection.Properties?.Title);
        Assert.Equal("segunda", inspection.Properties?.Keywords);
    }

    [Fact]
    public void CountsThePreviousVersionsEachUpdateLeaves()
    {
        var original = PdfFixtures.Classic();
        var once = Saved(original, new PropertiesDto(Title: "Primeira"));
        var twice = Saved(once, new PropertiesDto(Keywords: "segunda"));

        Assert.Equal(0, PdfMetadata.Inspect(original).Updates);
        Assert.Equal(1, PdfMetadata.Inspect(once).Updates);
        Assert.Equal(2, PdfMetadata.Inspect(twice).Updates);
        Assert.Equal(1, PdfMetadata.Inspect(Saved(PdfFixtures.WithXrefStream(), new PropertiesDto(Title: "X"))).Updates);
    }

    [Fact]
    public void SecondIdentifierChangesAndTheFirstStays()
    {
        var saved = Text(Saved(PdfFixtures.Classic(), new PropertiesDto(Title: "Outro")));

        var trailer = saved[saved.LastIndexOf("trailer", StringComparison.Ordinal)..];
        Assert.Contains("/ID [<0123456789ABCDEF0123456789ABCDEF> <", trailer);
        Assert.DoesNotContain("<0123456789ABCDEF0123456789ABCDEF> <0123456789ABCDEF0123456789ABCDEF>", trailer);
    }

    [Fact]
    public void NonAsciiTextGoesInUtf16()
    {
        var saved = Saved(PdfFixtures.Classic(), new PropertiesDto(Title: "Ação — 2026"));

        Assert.Contains("<FEFF", Text(saved[PdfFixtures.Classic().Length..]));
        Assert.Equal("Ação — 2026", PdfMetadata.Inspect(saved).Properties?.Title);
    }

    [Fact]
    public void EncryptedPdfIsRejected()
    {
        var encrypted = PdfFixtures.Classic(trailerExtras: " /Encrypt << /Filter /Standard /V 2 >>");

        Assert.True(PdfMetadata.Inspect(encrypted).Encrypted);
        Assert.Null(PdfMetadata.Inspect(encrypted).Properties);
        var refusal = Assert.Throws<PdfException>(() => PdfMetadata.Update(encrypted, new PropertiesDto(Title: "X")));
        Assert.Equal(PdfMetadata.EncryptedRefusal, refusal.Message);
    }

    [Fact]
    public void RecognizesASignedPdf()
    {
        var signed = PdfFixtures.Classic(catalogExtras: " /AcroForm << /Fields [] /SigFlags 3 >>");

        Assert.True(PdfMetadata.Inspect(signed).Signed);
    }

    [Fact]
    public void FileThatIsNotAPdfIsRejected()
    {
        Assert.Throws<PdfException>(() => PdfMetadata.Inspect(Encoding.ASCII.GetBytes("PK\u0003\u0004 não é PDF")));
    }

    [Fact]
    public void PdfWithoutCrossReferenceTableIsRejected()
    {
        var broken = Encoding.ASCII.GetBytes("%PDF-1.4\n1 0 obj\n<< >>\nendobj\n");

        Assert.Throws<PdfException>(() => PdfMetadata.Inspect(broken));
    }
}

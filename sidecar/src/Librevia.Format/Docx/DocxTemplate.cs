using System.IO.Compression;
using System.Text;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>The <c>docx.create</c> request: the document's page and styles.</summary>
public sealed record DocxCreateDto(
    [property: JsonPropertyName("page")] PageSetupDto Page,
    [property: JsonPropertyName("styles")] StyleSheetDto? Styles = null);

/// <summary>
/// The package that plays the original for a document born in the editor: without <c>oid</c>, every
/// block is written as new. The minimum: no theme (the font goes in <c>docDefaults</c>), no
/// <c>numbering.xml</c> (<see cref="NumberingFactory"/> creates it) and no author in
/// <c>docProps/core.xml</c>.
/// </summary>
public static class DocxTemplate
{
    /// <summary>The earliest zip date, so the same request gives the same bytes.</summary>
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    public static byte[] Create(PageSetupDto page, StyleSheetDto? styles = null)
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            // Fixed ids: the ones the SDK draws at random would change the bytes.
            var main = document.AddNewPart<MainDocumentPart>(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>("rId1");
            stylesPart.Styles = TemplateStyles.Create(styles);

            var settings = main.AddNewPart<DocumentSettingsPart>("rId2");
            settings.Settings = Settings();

            var fonts = main.AddNewPart<FontTablePart>("rId3");
            fonts.Fonts = FontTable();

            var section = Section(page);
            PlainBandWriter.Apply(main, section, page);

            main.Document = new Document(new Body(section));

            var core = document.AddNewPart<CoreFilePropertiesPart>("rId2");
            using (var stream = core.GetStream(FileMode.Create))
            {
                var xml = Encoding.UTF8.GetBytes(CoreXml);
                stream.Write(xml, 0, xml.Length);
            }

            var app = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            app.Properties = new Properties(new Application("Librevia"));
        }

        return WithFixedDates(buffer.ToArray());
    }

    /// <summary>
    /// By the same math as saving (<c>DocxWriter.ApplyPageSetup</c>); header, footer and gutter,
    /// which it does not write, are required by the schema.
    /// </summary>
    private static SectionProperties Section(PageSetupDto page)
    {
        var section = new SectionProperties(
            new DocumentFormat.OpenXml.Wordprocessing.PageSize(),
            new PageMargin
            {
                Header = (uint)Math.Max(0, Attr.MmToTwips(page.HeaderDistanceMm)),
                Footer = (uint)Math.Max(0, Attr.MmToTwips(page.FooterDistanceMm)),
                Gutter = 0U,
            });
        DocxWriter.ApplyPageSetup(section, page);
        return section;
    }

    /// <summary>
    /// Without <c>compatibilityMode</c> 15 Word opens in compatibility mode, with 2007 rules. The
    /// default tab stop is Portuguese Word's, 1.25 cm.
    /// </summary>
    private static Settings Settings() => new(
        new DefaultTabStop { Val = 708 },
        new CharacterSpacingControl { Val = CharacterSpacingValues.DoNotCompress },
        new Compatibility(new CompatibilitySetting
        {
            Name = CompatSettingNameValues.CompatibilityMode,
            Uri = "http://schemas.microsoft.com/office/word",
            Val = "15",
        }));

    private static Fonts FontTable() => new(
        Font(TemplateStyles.BodyFont, FontFamilyValues.Roman, "02020603050405020304"),
        Font(TemplateStyles.BandFont, FontFamilyValues.Swiss, "020F0502020204030204"));

    /// <c>panose</c> lets another program find the substitute, like Carlito for Calibri.
    private static Font Font(string name, FontFamilyValues family, string panose) => new(
        new Panose1Number { Val = panose },
        new FontCharSet { Val = "00" },
        new FontFamily { Val = family },
        new Pitch { Val = FontPitchValues.Variable })
    {
        Name = name,
    };

    /// <summary>By hand: the SDK has no typed DOM for this part. No author, on purpose.</summary>
    private const string CoreXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
        "<dc:creator></dc:creator></cp:coreProperties>";

    /// <summary>Repacks the zip with the same date on every entry.</summary>
    private static byte[] WithFixedDates(byte[] package)
    {
        using var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var result = new MemoryStream();

        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in source.Entries)
            {
                var copy = output.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                copy.LastWriteTime = ZipEpoch;

                using var from = entry.Open();
                using var to = copy.Open();
                from.CopyTo(to);
            }
        }

        return result.ToArray();
    }
}

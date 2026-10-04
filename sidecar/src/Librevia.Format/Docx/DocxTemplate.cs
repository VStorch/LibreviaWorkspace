using System.IO.Compression;
using System.Text;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>O pedido de <c>docx.create</c>: a página e os estilos do documento.</summary>
public sealed record DocxCreateDto(
    [property: JsonPropertyName("page")] PageSetupDto Page,
    [property: JsonPropertyName("styles")] StyleSheetDto? Styles = null);

/// <summary>
/// O pacote que faz o papel de original para o documento nascido no editor: sem
/// <c>oid</c>, todo bloco é gravado como novo. O mínimo: sem tema (a fonte vai nos
/// <c>docDefaults</c>), sem <c>numbering.xml</c> (o <see cref="NumberingFactory"/>
/// o cria) e sem autor em <c>docProps/core.xml</c>.
/// </summary>
public static class DocxTemplate
{
    /// <summary>A menor data do zip, para o mesmo pedido dar os mesmos bytes.</summary>
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    public static byte[] Create(PageSetupDto page, StyleSheetDto? styles = null)
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            // Ids fixos: os que o SDK sorteia mudariam os bytes.
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
    /// Pela mesma conta da gravação (<c>DocxWriter.ApplyPageSetup</c>); cabeçalho,
    /// rodapé e medianiz, que ela não escreve, o esquema exige.
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
    /// Sem <c>compatibilityMode</c> 15 o Word abre em modo de compatibilidade, com as
    /// regras de 2007. A tabulação padrão é a do Word em português, 1,25 cm.
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

    /// <summary>O <c>panose</c> deixa outro programa achar a substituta, como a Carlito para a Calibri.</summary>
    private static Font Font(string name, FontFamilyValues family, string panose) => new(
        new Panose1Number { Val = panose },
        new FontCharSet { Val = "00" },
        new FontFamily { Val = family },
        new Pitch { Val = FontPitchValues.Variable })
    {
        Name = name,
    };

    /// <summary>À mão: o SDK não tem DOM tipado para esta parte. Sem autor, de propósito.</summary>
    private const string CoreXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
        "<dc:creator></dc:creator></cp:coreProperties>";

    /// <summary>Reembala o zip com a mesma data em toda entrada.</summary>
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

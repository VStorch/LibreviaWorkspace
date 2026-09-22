using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O pacote de um documento que nasceu aqui dentro.
/// </summary>
/// <remarks>
/// A gravação cirúrgica pressupõe um original: ela parte dos bytes do arquivo e
/// troca só o que foi editado. Um documento novo não tem arquivo, e por isso não
/// podia ser salvo em DOCX. Este pacote faz o papel do original — o documento
/// segue pelo <see cref="DocxWriter"/> de sempre, e como nenhum bloco tem `oid`,
/// todos são gravados como novos.
///
/// Escrito em código, e não embutido como binário, por duas razões: o mesmo
/// pedido dá os mesmos bytes (as datas do zip são fixas e o `core.xml` não leva
/// data nenhuma), e cada parte passa pelo validador do SDK nos testes. Um `.docx`
/// guardado nos recursos seria uma caixa-preta que ninguém revisa.
///
/// O mínimo, e só o mínimo: sem tema (a fonte vai por nome nos `docDefaults`),
/// sem `numbering.xml` (o <see cref="NumberingFactory"/> o cria quando a primeira
/// lista pede) e sem nome de autor em `docProps/core.xml` — o arquivo não carrega
/// quem o escreveu a não ser que a pessoa o diga.
/// </remarks>
public static class DocxTemplate
{
    /// <summary>
    /// A data de toda entrada do zip: a menor que o formato aceita.
    /// </summary>
    /// <remarks>
    /// Com a hora da gravação, dois pacotes iguais saíam com bytes diferentes, e
    /// o teste de determinismo não teria como distinguir uma mudança de conteúdo
    /// de uma mudança de relógio.
    /// </remarks>
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    public static byte[] Create(PageSetupDto page)
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            // Ids fixos em toda relação: os que o SDK sorteia mudariam os bytes
            // a cada pacote.
            var main = document.AddNewPart<MainDocumentPart>(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");

            var styles = main.AddNewPart<StyleDefinitionsPart>("rId1");
            styles.Styles = TemplateStyles.Create();

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
    /// O `w:sectPr` do corpo: papel, margens e distâncias da faixa.
    /// </summary>
    /// <remarks>
    /// Pelo mesmo caminho da gravação (<c>DocxWriter.ApplyPageSetup</c>), para
    /// que o papel do pacote novo e o do pacote regravado sejam a mesma conta. Os
    /// três atributos que ele não escreve — cabeçalho, rodapé e medianiz — são
    /// obrigatórios no esquema e entram aqui.
    /// </remarks>
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

    /// <remarks>
    /// `compatibilityMode` 15 é o Word 2013 em diante: sem ele o Word abre o
    /// arquivo em modo de compatibilidade e desenha com as regras de 2007. A
    /// parada de tabulação padrão é a do Word em português, 1,25 cm.
    /// </remarks>
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

    /// <remarks>
    /// O `panose` é o que deixa outro programa escolher a substituta certa quando
    /// a fonte falta — a Liberation Serif no lugar da Times, a Carlito no lugar
    /// da Calibri.
    /// </remarks>
    private static Font Font(string name, FontFamilyValues family, string panose) => new(
        new Panose1Number { Val = panose },
        new FontCharSet { Val = "00" },
        new FontFamily { Val = family },
        new Pitch { Val = FontPitchValues.Variable })
    {
        Name = name,
    };

    /// <remarks>
    /// Escrito à mão porque o SDK não tem DOM tipado para esta parte. O autor
    /// fica vazio de propósito; o Word preenche o dele ao salvar, se a pessoa
    /// quiser.
    /// </remarks>
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

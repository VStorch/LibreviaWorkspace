using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// <summary>
/// As propriedades do documento (M11): `docProps/core.xml` e `docProps/app.xml`
/// lidas, e gravadas só quando algum campo mudou — com o resto da parte intacto.
/// </summary>
public class PropertiesTests
{
    private const string Core =
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""" + "\r\n" +
        """<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">""" +
        """<dc:title>Relatório anual</dc:title><dc:subject>Contas</dc:subject><dc:creator>Ana Lima</dc:creator>""" +
        """<cp:keywords>contas; 2025</cp:keywords><dc:description>Para a assembleia.</dc:description>""" +
        """<cp:lastModifiedBy>Bia</cp:lastModifiedBy><cp:revision>7</cp:revision>""" +
        """<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-02T03:04:05Z</dcterms:created>""" +
        """<dcterms:modified xsi:type="dcterms:W3CDTF">2025-02-03T04:05:06Z</dcterms:modified>""" +
        """<cp:category>Finanças</cp:category><cp:contentStatus>Rascunho</cp:contentStatus></cp:coreProperties>""";

    private const string App =
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""" + "\r\n" +
        """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">""" +
        """<Template>Normal.dotm</Template><TotalTime>42</TotalTime><Pages>1</Pages><Application>Microsoft Office Word</Application>""" +
        """<Manager>Carla</Manager><Company>ACME</Company></Properties>""";

    private const string Custom =
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""" + "\r\n" +
        """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Cliente"><vt:lpwstr>XPTO</vt:lpwstr></property></Properties>""";

    /// <summary>O documento simples com as três partes de propriedades, como o Word as grava.</summary>
    private static byte[] WithProperties()
    {
        using var buffer = new MemoryStream();
        var simple = Fixtures.Simple();
        buffer.Write(simple, 0, simple.Length);
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            Write(document.CoreFilePropertiesPart ?? document.AddCoreFilePropertiesPart(), Core);
            Write(document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart(), App);
            Write(document.CustomFilePropertiesPart ?? document.AddCustomFilePropertiesPart(), Custom);
        }

        return buffer.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create);
        var bytes = Encoding.UTF8.GetBytes(xml);
        stream.Write(bytes, 0, bytes.Length);
    }

    [Fact]
    public void LeAsPropriedadesDoCoreEDoApp()
    {
        var properties = Open(WithProperties()).Properties;

        Assert.NotNull(properties);
        Assert.Equal("Relatório anual", properties.Title);
        Assert.Equal("Contas", properties.Subject);
        Assert.Equal("Ana Lima", properties.Creator);
        Assert.Equal("contas; 2025", properties.Keywords);
        Assert.Equal("Finanças", properties.Category);
        Assert.Equal("Para a assembleia.", properties.Description);
        Assert.Equal("Bia", properties.LastModifiedBy);
        Assert.Equal("7", properties.Revision);
        Assert.Equal("2025-01-02T03:04:05Z", properties.Created);
        Assert.Equal("2025-02-03T04:05:06Z", properties.Modified);
        Assert.Equal("ACME", properties.Company);
        Assert.Equal("Carla", properties.Manager);
        Assert.Equal(42, properties.TotalTime);
    }

    [Fact]
    public void GravarSemMudarDeixaAsPartesDePropriedadesByteAByte()
    {
        var original = WithProperties();
        var before = PartsOf(original);

        // O modelo como a leitura o deu, propriedades incluídas, e também sem elas.
        foreach (var model in new[] { Clone(Open(original)), Clone(Open(original)) with { Properties = null } })
        {
            var after = PartsOf(Save(original, model).Bytes);
            foreach (var part in new[] { "docProps/core.xml", "docProps/app.xml", "docProps/custom.xml" })
            {
                Assert.Equal(before[part], after[part]);
            }
        }
    }

    [Fact]
    public void MudarOTituloRegravaSoOCoreEGuardaOQueNaoConhece()
    {
        var original = WithProperties();
        var model = Clone(Open(original));
        model = model with { Properties = model.Properties! with { Title = "Novo título", Keywords = "a; b" } };

        var (bytes, result) = Save(original, model);
        var before = PartsOf(original);
        var after = PartsOf(bytes);

        Assert.Equal(0, result.RewrittenBlocks);
        Assert.NotEqual(before["docProps/core.xml"], after["docProps/core.xml"]);
        Assert.Equal(before["docProps/app.xml"], after["docProps/app.xml"]);
        Assert.Equal(before["docProps/custom.xml"], after["docProps/custom.xml"]);
        Assert.Equal(before["word/document.xml"], after["word/document.xml"]);

        var core = XmlOf(bytes, "docProps/core.xml");
        Assert.Contains("<cp:contentStatus>Rascunho</cp:contentStatus>", core, StringComparison.Ordinal);
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", core, StringComparison.Ordinal);

        var reread = Open(bytes).Properties!;
        Assert.Equal("Novo título", reread.Title);
        Assert.Equal("a; b", reread.Keywords);
        Assert.Equal("Ana Lima", reread.Creator);
        Assert.Equal("2025-02-03T04:05:06Z", reread.Modified);
    }

    [Fact]
    public void ACadeiaVaziaApagaOCampoEAEmpresaRegravaSoOApp()
    {
        var original = WithProperties();
        var model = Clone(Open(original)) with
        {
            Properties = new PropertiesDto(Subject: string.Empty, Company: "Outra"),
        };

        var bytes = Save(original, model).Bytes;
        var reread = Open(bytes).Properties!;

        Assert.Null(reread.Subject);
        Assert.Equal("Relatório anual", reread.Title);
        Assert.Equal("Outra", reread.Company);
        Assert.Equal("Carla", reread.Manager);
        Assert.Equal(42, reread.TotalTime);
        Assert.Contains("<Template>Normal.dotm</Template>", XmlOf(bytes, "docProps/app.xml"), StringComparison.Ordinal);
        Assert.Equal(PartsOf(original)["docProps/custom.xml"], PartsOf(bytes)["docProps/custom.xml"]);
    }

    [Fact]
    public void ODocumentoNovoGanhaCriadorEDataNoCore()
    {
        var page = Open(Fixtures.Simple()).Page;
        var template = DocxTemplate.Create(page);
        var model = Clone(Open(template)) with
        {
            Properties = new PropertiesDto(
                Creator: "Ana Lima",
                Created: "2026-10-02T12:00:00Z",
                Modified: "2026-10-02T12:00:00Z",
                LastModifiedBy: "Ana Lima",
                Revision: "1"),
        };

        var bytes = Save(template, model).Bytes;
        var core = XmlOf(bytes, "docProps/core.xml");

        Assert.Contains("<dc:creator>Ana Lima</dc:creator>", core, StringComparison.Ordinal);
        Assert.Contains(
            """<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-02T12:00:00Z</dcterms:created>""",
            core,
            StringComparison.Ordinal);
        var reread = Open(bytes).Properties!;
        Assert.Equal("Ana Lima", reread.LastModifiedBy);
        Assert.Equal("1", reread.Revision);
    }

    [Fact]
    public void OPacoteSemCorePassaATerUm()
    {
        var original = Fixtures.Simple();
        Assert.False(PartsOf(original).ContainsKey("docProps/core.xml"));

        var model = Clone(Open(original)) with { Properties = new PropertiesDto(Title: "Com título") };
        var bytes = Save(original, model).Bytes;

        Assert.Equal("Com título", Open(bytes).Properties!.Title);
    }
}

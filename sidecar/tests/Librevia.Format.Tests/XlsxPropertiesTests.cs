using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;
using Librevia.Format.Xlsx;

namespace Librevia.Format.Tests;

/// The same <c>docProps/</c> parts as a document's, read and patched by the same code.
public class XlsxPropertiesTests
{
    /// <summary>The sales spreadsheet with the properties Excel writes.</summary>
    private static byte[] WithProperties()
    {
        using var buffer = new MemoryStream();
        buffer.Write(XlsxFixtures.Sales());
        using (var package = SpreadsheetDocument.Open(buffer, true))
        {
            var core = package.CoreFilePropertiesPart ?? package.AddCoreFilePropertiesPart();
            var app = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
            PropertiesTests.Write(core, PropertiesTests.Core);
            PropertiesTests.Write(app, PropertiesTests.App);
        }

        return buffer.ToArray();
    }

    private static PropertiesDto? PropertiesOf(byte[] bytes) => XlsxReader.Read(bytes).Workbook.Properties;

    private static byte[] Save(byte[]? original, PropertiesDto? properties)
    {
        var model = original is null
            ? new WorkbookDto { Sheets = { new SheetDto { Name = "Plan1" } } }
            : XlsxReader.Read(original).Workbook;
        model.Properties = properties;
        return XlsxWriter.Write(original, model).Bytes;
    }

    [Fact]
    public void ReadsTheSpreadsheetProperties()
    {
        var properties = PropertiesOf(WithProperties());

        Assert.NotNull(properties);
        Assert.Equal("Relatório anual", properties.Title);
        Assert.Equal("Ana Lima", properties.Creator);
        Assert.Equal("7", properties.Revision);
        Assert.Equal("ACME", properties.Company);
        Assert.Equal("Carla", properties.Manager);
    }

    [Fact]
    public void SavingWithoutPropertiesKeepsTheOnesInTheFile()
    {
        var saved = PropertiesOf(Save(WithProperties(), properties: null));

        Assert.NotNull(saved);
        Assert.Equal("Relatório anual", saved.Title);
        Assert.Equal("Bia", saved.LastModifiedBy);
        Assert.Equal("7", saved.Revision);
        Assert.Equal("ACME", saved.Company);
    }

    [Fact]
    public void WritesOnlyTheFieldsThatChanged()
    {
        var patch = new PropertiesDto(Title: "Orçamento", Subject: "", Revision: "8");
        var saved = PropertiesOf(Save(WithProperties(), patch));

        Assert.NotNull(saved);
        Assert.Equal("Orçamento", saved.Title);
        Assert.Null(saved.Subject);
        Assert.Equal("8", saved.Revision);
        Assert.Equal("Ana Lima", saved.Creator);
        Assert.Equal("Carla", saved.Manager);
    }

    [Fact]
    public void NewSpreadsheetGetsTheProperties()
    {
        var saved = PropertiesOf(Save(null, new PropertiesDto(Title: "Nova", Creator: "Ana", Company: "ACME")));

        Assert.NotNull(saved);
        Assert.Equal("Nova", saved.Title);
        Assert.Equal("Ana", saved.Creator);
        Assert.Equal("ACME", saved.Company);
    }
}

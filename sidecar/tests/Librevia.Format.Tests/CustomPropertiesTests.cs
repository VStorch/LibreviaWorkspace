using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Librevia.Format.Docx;
using Librevia.Format.Xlsx;
using static Librevia.Format.Tests.Roundtrip;

namespace Librevia.Format.Tests;

/// File → Properties → Custom: read from <c>docProps/custom.xml</c>, and written only when the list
/// changed, keeping what the editor does not show.
public class CustomPropertiesTests
{
    private const string Part = "docProps/custom.xml";

    private const string Custom =
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""" + "\r\n" +
        """<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Cliente"><vt:lpwstr>XPTO</vt:lpwstr></property>""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Processo"><vt:i8>2026</vt:i8></property>""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Valor"><vt:r8>12.5</vt:r8></property>""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Prazo"><vt:filetime>2026-10-12T03:00:00Z</vt:filetime></property>""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="Aprovado"><vt:bool>1</vt:bool></property>""" +
        """<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Revisores"><vt:vector size="1" baseType="lpwstr"><vt:lpwstr>Ana</vt:lpwstr></vt:vector></property>""" +
        """</Properties>""";

    private static byte[] WithCustom()
    {
        using var buffer = new MemoryStream();
        buffer.Write(Fixtures.Simple());
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            PropertiesTests.Write(document.AddCustomFilePropertiesPart(), Custom);
        }

        return buffer.ToArray();
    }

    private static byte[] SaveCustom(byte[] original, List<CustomPropertyDto> custom)
    {
        var model = Open(original);
        return Save(original, model with { Properties = (model.Properties ?? new PropertiesDto()) with { Custom = custom } }).Bytes;
    }

    private static List<CustomPropertyDto> Shown(byte[] docx) => Open(docx).Properties?.Custom ?? [];

    [Fact]
    public void LeCadaTipoEDeixaDeForaOQueOEditorNaoMostra()
    {
        Assert.Equal(
            [
                new CustomPropertyDto("Cliente", "text", "XPTO"),
                new CustomPropertyDto("Processo", "number", "2026"),
                new CustomPropertyDto("Valor", "number", "12.5"),
                new CustomPropertyDto("Prazo", "date", "2026-10-12T03:00:00Z"),
                new CustomPropertyDto("Aprovado", "boolean", "true"),
            ],
            Shown(WithCustom()));
    }

    [Fact]
    public void AMesmaListaEmOutraOrdemNaoReescreveAParte()
    {
        var original = WithCustom();
        var reversed = Shown(original).AsEnumerable().Reverse().ToList();

        var saved = SaveCustom(original, reversed);

        Assert.Equal(Custom, XmlOf(saved, Part));
    }

    [Fact]
    public void AlterarUmValorTrocaSoEleEGuardaOVetor()
    {
        var original = WithCustom();
        var custom = Shown(original).Select(property =>
            property.Name == "Processo" ? property with { Value = "2027" } : property).ToList();

        var xml = XmlOf(SaveCustom(original, custom), Part);

        Assert.Contains("""pid="3" name="Processo"><vt:i8>2027</vt:i8>""", xml);
        Assert.Contains("<vt:lpwstr>XPTO</vt:lpwstr>", xml);
        Assert.Contains("""<vt:vector size="1" baseType="lpwstr">""", xml);
    }

    [Fact]
    public void ApagarETrocarOTipoMantemOsPids()
    {
        var original = WithCustom();
        var custom = Shown(original)
            .Where(property => property.Name != "Cliente")
            .Select(property => property.Name == "Valor" ? new CustomPropertyDto("Valor", "text", "doze") : property)
            .Append(new CustomPropertyDto("Setor", "text", "Jurídico"))
            .ToList();

        var saved = SaveCustom(original, custom);

        var xml = XmlOf(saved, Part);
        Assert.DoesNotContain("Cliente", xml);
        Assert.Contains("""pid="4" name="Valor"><vt:lpwstr>doze</vt:lpwstr>""", xml);
        Assert.Contains("""fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="8" name="Setor"><vt:lpwstr>Jurídico</vt:lpwstr>""", xml);
        Assert.Equal(custom.OrderBy(property => property.Name), Shown(saved).OrderBy(property => property.Name));
        AssertSchema(saved);
    }

    [Fact]
    public void ListaVaziaSemNadaDesconhecidoTiraAParte()
    {
        using var buffer = new MemoryStream();
        buffer.Write(Fixtures.Simple());
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            PropertiesTests.Write(document.AddCustomFilePropertiesPart(), PropertiesTests.Custom);
        }

        var saved = SaveCustom(buffer.ToArray(), []);

        Assert.DoesNotContain(Part, PartsOf(saved).Keys);
        Assert.DoesNotContain("custom", XmlOf(saved, "_rels/.rels"));
        Assert.DoesNotContain("custom", XmlOf(saved, "[Content_Types].xml"));
        AssertSchema(saved);
    }

    [Fact]
    public void DocumentoSemAParteGanhaUmaComONumeroEAData()
    {
        var saved = SaveCustom(Fixtures.Simple(),
        [
            new CustomPropertyDto("Folhas", "number", "3"),
            new CustomPropertyDto("Entrega", "date", "2026-10-12T03:00:00Z"),
            new CustomPropertyDto("Urgente", "boolean", "false"),
        ]);

        var xml = XmlOf(saved, Part);
        Assert.Contains("""pid="2" name="Folhas"><vt:i4>3</vt:i4>""", xml);
        Assert.Contains("<vt:filetime>2026-10-12T03:00:00Z</vt:filetime>", xml);
        Assert.Contains("<vt:bool>false</vt:bool>", xml);
        Assert.Equal(3, Shown(saved).Count);
        AssertSchema(saved);
    }

    [Fact]
    public void NaPlanilhaAsPersonalizadasVaoEVoltam()
    {
        using var book = new XLWorkbook();
        book.Worksheets.Add("Plan1").Cell("A1").Value = 1;
        using var stream = new MemoryStream();
        book.SaveAs(stream);
        var original = stream.ToArray();

        var model = XlsxReader.Read(original).Workbook;
        model.Properties = new PropertiesDto(Custom: [new CustomPropertyDto("Cliente", "text", "XPTO")]);
        var saved = XlsxWriter.Write(original, model).Bytes;

        Assert.Equal([new CustomPropertyDto("Cliente", "text", "XPTO")], XlsxReader.Read(saved).Workbook.Properties?.Custom);
    }
}

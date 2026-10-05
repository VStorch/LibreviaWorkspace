using ClosedXML.Excel;
using Librevia.Format.Xlsx;

namespace Librevia.Format.Tests;

/// <summary>Test spreadsheets built in code, readable in review.</summary>
public static class XlsxFixtures
{
    /// <summary>
    /// Sales with a formula, currency, a date and a second tab. The 14 pt font in B2 is what the
    /// model does not represent, and proves preservation.
    /// </summary>
    public static byte[] Sales()
    {
        using var book = new XLWorkbook();

        var sheet = book.Worksheets.Add("Vendas");
        sheet.Cell("A1").Value = "Produto";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Cell("B1").Value = "Qtd";
        sheet.Cell("C1").Value = "Preço";
        sheet.Cell("D1").Value = "Total";

        sheet.Cell("A2").Value = "Cabo";
        sheet.Cell("B2").Value = 3;
        sheet.Cell("B2").Style.Font.FontSize = 14;
        sheet.Cell("C2").Value = 12.5;
        sheet.Cell("C2").Style.NumberFormat.Format = "\"R$\" #,##0.00";
        sheet.Cell("D2").FormulaA1 = "B2*C2";
        sheet.Cell("D2").Style.NumberFormat.Format = "\"R$\" #,##0.00";

        sheet.Cell("A3").Value = "Fonte";
        sheet.Cell("B3").Value = 2;
        sheet.Cell("C3").Value = 89.9;
        sheet.Cell("D3").FormulaA1 = "B3*C3";

        sheet.Cell("E2").Value = new DateTime(2026, 3, 15);
        sheet.Cell("E2").Style.NumberFormat.NumberFormatId = 14;

        sheet.Cell("D5").FormulaA1 = "SUM(D2:D3)";
        sheet.Cell("A6").FormulaA1 = "Resumo!B1";

        sheet.Column(1).Width = 20;
        sheet.SheetView.Freeze(1, 0);

        var summary = book.Worksheets.Add("Resumo");
        summary.Cell("A1").Value = "Itens";
        summary.Cell("B1").Value = 7;

        using var stream = new MemoryStream();
        book.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// What the model does not represent and saving preserves: filter, merge, conditional
    /// formatting and data validation.
    /// </summary>
    public static byte[] WithUnmodeledFeatures()
    {
        using var book = new XLWorkbook();
        var sheet = book.Worksheets.Add("Dados");

        sheet.Cell("A1").Value = "Setor";
        sheet.Cell("B1").Value = "Gasto";
        sheet.Cell("A2").Value = "Compras";
        sheet.Cell("B2").Value = 1200;
        sheet.Cell("A3").Value = "TI";
        sheet.Cell("B3").Value = 8400;

        sheet.Range("A1:B3").SetAutoFilter();
        sheet.Range("D1:F1").Merge();
        sheet.Range("B2:B3").AddConditionalFormat().WhenGreaterThan(5000).Fill.SetBackgroundColor(XLColor.Red);
        sheet.Range("A2:A3").CreateDataValidation().List("\"Compras,TI\"");

        using var stream = new MemoryStream();
        book.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>A sheet with merged cells, so the inventory has something to say.</summary>
    public static byte[] WithMerge()
    {
        using var book = new XLWorkbook();
        var sheet = book.Worksheets.Add("Plan1");
        sheet.Cell("A1").Value = "Título largo";
        sheet.Range("A1:C1").Merge();

        using var stream = new MemoryStream();
        book.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// A part ClosedXML does not model: if it survives, charts and pivot tables, which also have
    /// their own parts, survive.
    /// </summary>
    public static byte[] WithForeignPart(byte[] original, string content)
    {
        using var source = new MemoryStream(original, writable: false);
        using var target = new MemoryStream();
        source.CopyTo(target);
        target.Position = 0;

        using (var zip = new System.IO.Compression.ZipArchive(
                   target, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.CreateEntry("customXml/item1.xml");
            using var writer = new StreamWriter(entry.Open());
            writer.Write(content);
        }

        return target.ToArray();
    }

    public static string? PartText(byte[] package, string path)
    {
        using var stream = new MemoryStream(package, writable: false);
        using var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read);
        var entry = zip.GetEntry(path);
        if (entry is null) return null;

        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Reads a package and returns only the model, which is what the tests look at.
    /// </summary>
    public static WorkbookDto Model(byte[] bytes) => XlsxReader.Read(bytes).Workbook;

    public static SheetDto Sheet(WorkbookDto workbook, string name) =>
        workbook.Sheets.Single(sheet => sheet.Name == name);
}

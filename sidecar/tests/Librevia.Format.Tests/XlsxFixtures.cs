using ClosedXML.Excel;
using Librevia.Format.Xlsx;

namespace Librevia.Format.Tests;

/// <summary>
/// Planilhas de teste construídas em código, pelo mesmo motivo dos fixtures de
/// documento: o que elas contêm fica legível na revisão, e um binário no git é
/// uma caixa preta que ninguém confere.
/// </summary>
public static class XlsxFixtures
{
    /// <summary>
    /// Planilha de vendas com fórmula, moeda, data e uma segunda aba.
    /// </summary>
    /// <remarks>
    /// A fonte de 14 pontos em B2 não é enfeite: é o atributo que o modelo do
    /// aplicativo **não** representa, e é ele que prova que a gravação
    /// cirúrgica preserva o que não foi editado.
    /// </remarks>
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

    /// <summary>Planilha com células mescladas, para o inventário ter o que dizer.</summary>
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
    /// Acrescenta ao pacote uma parte que o ClosedXML não modela.
    /// </summary>
    /// <remarks>
    /// É o teste que separa "preserva" de "regenera": se esta parte sobrevive a
    /// uma gravação, gráficos e tabelas dinâmicas — que também vivem em partes
    /// próprias — sobrevivem pelo mesmo mecanismo.
    /// </remarks>
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

    /// <summary>Lê um pacote e devolve só o modelo, que é o que os testes olham.</summary>
    public static WorkbookDto Model(byte[] bytes) => XlsxReader.Read(bytes).Workbook;

    public static SheetDto Sheet(WorkbookDto workbook, string name) =>
        workbook.Sheets.Single(sheet => sheet.Name == name);
}

using ClosedXML.Excel;
using Librevia.Format.Xlsx;
using Xunit;

namespace Librevia.Format.Tests;

public class XlsxReadTests
{
    [Fact]
    public void LeValoresDeCadaTipo()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        Assert.Equal("Produto", sheet.Cells["A1"].Value);
        Assert.Equal(3d, Assert.IsType<double>(sheet.Cells["B2"].Value));
        Assert.Equal(12.5d, Assert.IsType<double>(sheet.Cells["C2"].Value));
    }

    [Fact]
    public void LeFormulaComOSinalDeIgual()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        // O ClosedXML entrega sem o `=`; o modelo do aplicativo guarda com.
        Assert.Equal("=B2*C2", sheet.Cells["D2"].Formula);
        Assert.Equal("=SUM(D2:D3)", sheet.Cells["D5"].Formula);
    }

    [Fact]
    public void LeFormulaEntreAbas()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        Assert.Equal("=Resumo!B1", sheet.Cells["A6"].Formula);
    }

    [Fact]
    public void DataViraNumeroDeSerie()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");
        var cell = sheet.Cells["E2"];

        // 15/03/2026 na contagem que começa em 1899-12-30.
        Assert.Equal(46096d, Assert.IsType<double>(cell.Value));
        Assert.Equal("date", cell.Style!.Format);
    }

    [Fact]
    public void ReconheceMoedaPelaMascara()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        Assert.Equal("currency", sheet.Cells["C2"].Style!.Format);
        Assert.Equal(2, sheet.Cells["C2"].Style!.Decimals);
    }

    [Fact]
    public void LeNegrito()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        Assert.True(sheet.Cells["A1"].Style!.Bold);
    }

    [Fact]
    public void LeLarguraDeColunaEmPixels()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        // 20 caracteres na fonte padrão dão 145 pixels.
        Assert.Equal(145d, sheet.ColumnWidths[0]);
    }

    [Fact]
    public void LeCongelamento()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        Assert.Equal(1, sheet.FrozenRows);
        Assert.Equal(0, sheet.FrozenColumns);
    }

    [Fact]
    public void LeTodasAsAbasNaOrdem()
    {
        var workbook = XlsxFixtures.Model(XlsxFixtures.Sales());

        Assert.Equal(["Vendas", "Resumo"], workbook.Sheets.Select(sheet => sheet.Name));
    }

    [Fact]
    public void CelulaVaziaNaoEntraNoMapa()
    {
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(XlsxFixtures.Sales()), "Vendas");

        // O mapa é esparso: guardar célula vazia faria o arquivo crescer à toa.
        Assert.False(sheet.Cells.ContainsKey("Z99"));
    }

    [Fact]
    public void ArquivoQueNaoEPlanilhaDaMensagemCompreensivel()
    {
        var problem = Assert.Throws<XlsxException>(() => XlsxReader.Read([1, 2, 3, 4]));

        Assert.Contains("não pôde ser lida", problem.Message);
    }
}

public class XlsxInventoryTests
{
    [Fact]
    public void AvisaSobreCelulasMescladas()
    {
        var result = XlsxReader.Read(XlsxFixtures.WithMerge());

        Assert.Contains(result.Inventory.Invisible, message => message.Contains("mescladas"));
    }

    [Fact]
    public void PlanilhaComumNaoGeraAviso()
    {
        // Um aviso que aparece em todo arquivo é um aviso que o usuário aprende
        // a fechar sem ler.
        var result = XlsxReader.Read(XlsxFixtures.Sales());

        // Comparar as frases, e não só `IsEmpty`, porque a falha precisa dizer
        // **qual** aviso apareceu — foi assim que o comentário fantasma do
        // ClosedXML se identificou.
        Assert.Equal(string.Empty, string.Join(" | ", result.Inventory.Invisible.Concat(result.Inventory.Lost)));
    }
}

public class XlsxSurgicalTests
{
    /// <summary>Grava o modelo tal como foi lido, sem editar nada.</summary>
    private static (byte[] Bytes, XlsxWriter.Result Report) SaveUnchanged(byte[] original)
    {
        var model = XlsxFixtures.Model(original);
        return XlsxWriter.Write(original, model);
    }

    [Fact]
    public void GravarSemEditarNaoEscreveNenhumaCelula()
    {
        // É a medida que dá sentido à palavra "cirúrgico": se gravar sem editar
        // já reescrevesse tudo, não haveria preservação nenhuma.
        var (_, report) = SaveUnchanged(XlsxFixtures.Sales());

        Assert.Equal(0, report.CellsWritten);
        Assert.Equal(0, report.CellsCleared);
        Assert.True(report.CellsPreserved > 10);
    }

    [Fact]
    public void ModeloQueVemPeloJsonTambemPreserva()
    {
        // O aplicativo não entrega um `WorkbookDto`: entrega JSON, e é essa a
        // única forma que chega aqui em produção. Sem a conversão de escalares,
        // cada valor voltaria como `JsonElement`, nenhuma célula pareceria igual
        // à do arquivo, e a gravação reescreveria a planilha inteira — apagando
        // justamente o que a cirurgia existe para preservar.
        var original = XlsxFixtures.Sales();
        var json = System.Text.Json.JsonSerializer.Serialize(
            XlsxFixtures.Model(original), Protocol.JsonOptions.Default);
        var model = System.Text.Json.JsonSerializer.Deserialize<WorkbookDto>(
            json, Protocol.JsonOptions.Default);

        var (_, report) = XlsxWriter.Write(original, model!);

        Assert.Equal(0, report.CellsWritten);
        Assert.Equal(0, report.CellsCleared);
        Assert.True(report.CellsPreserved > 10);
    }

    [Fact]
    public void EditarUmaCelulaEscreveApenasUma()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells["A2"].Value = "Cabo HDMI";

        var (_, report) = XlsxWriter.Write(original, model);

        Assert.Equal(1, report.CellsWritten);
    }

    [Fact]
    public void PreservaOQueOModeloNaoRepresenta()
    {
        // B2 tem fonte de 14 pontos, que o modelo do aplicativo não carrega.
        // Editar a célula ao lado não pode levá-la junto.
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells["A2"].Value = "Outro nome";

        var (bytes, _) = XlsxWriter.Write(original, model);

        using var stream = new MemoryStream(bytes, writable: false);
        using var book = new XLWorkbook(stream);
        Assert.Equal(14, book.Worksheet("Vendas").Cell("B2").Style.Font.FontSize);
    }

    [Fact]
    public void PreservaParteDoPacoteQueNaoConhece()
    {
        // Gráfico e tabela dinâmica vivem em partes próprias, e sobrevivem pelo
        // mesmo mecanismo que esta.
        var original = XlsxFixtures.WithForeignPart(XlsxFixtures.Sales(), "<catalogo>NAO-PERCA</catalogo>");
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells["A2"].Value = "Editado";

        var (bytes, _) = XlsxWriter.Write(original, model);

        Assert.Contains("NAO-PERCA", XlsxFixtures.PartText(bytes, "customXml/item1.xml"));
    }

    [Fact]
    public void ValorEditadoChegaNoArquivo()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells["A2"].Value = "Cabo HDMI 2m";

        var (bytes, _) = XlsxWriter.Write(original, model);

        Assert.Equal("Cabo HDMI 2m", XlsxFixtures.Sheet(XlsxFixtures.Model(bytes), "Vendas").Cells["A2"].Value);
    }

    [Fact]
    public void FormulaNovaChegaComoFormula()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells["D3"] = new CellDto { Formula = "=B3*C3*2" };

        var (bytes, _) = XlsxWriter.Write(original, model);

        Assert.Equal("=B3*C3*2", XlsxFixtures.Sheet(XlsxFixtures.Model(bytes), "Vendas").Cells["D3"].Formula);
    }

    [Fact]
    public void CelulaApagadaSaiDoArquivo()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells.Remove("A3");

        var (bytes, report) = XlsxWriter.Write(original, model);

        Assert.Equal(1, report.CellsCleared);
        Assert.False(XlsxFixtures.Sheet(XlsxFixtures.Model(bytes), "Vendas").Cells.ContainsKey("A3"));
    }

    [Fact]
    public void FormatoEscolhidoNoAplicativoChegaComoMascara()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        XlsxFixtures.Sheet(model, "Vendas").Cells["B3"].Style =
            new CellStyleDto { Format = "percent", Decimals = 1 };

        var (bytes, _) = XlsxWriter.Write(original, model);
        var style = XlsxFixtures.Sheet(XlsxFixtures.Model(bytes), "Vendas").Cells["B3"].Style;

        Assert.Equal("percent", style!.Format);
        Assert.Equal(1, style.Decimals);
    }

    [Fact]
    public void RenomearAbaVaiParaOArquivo()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        model.Sheets[1].Name = "Totais";

        var (bytes, _) = XlsxWriter.Write(original, model);

        Assert.Equal(["Vendas", "Totais"], XlsxFixtures.Model(bytes).Sheets.Select(sheet => sheet.Name));
    }

    [Fact]
    public void TrocarNomesEntreDuasAbasNaoColide()
    {
        // Dar a uma aba o nome que a outra ainda tem é erro do ClosedXML, e
        // trocar dois nomes é um caso perfeitamente normal.
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        model.Sheets[0].Name = "Resumo";
        model.Sheets[1].Name = "Vendas";

        var (bytes, _) = XlsxWriter.Write(original, model);

        Assert.Equal(["Resumo", "Vendas"], XlsxFixtures.Model(bytes).Sheets.Select(sheet => sheet.Name));
    }

    [Fact]
    public void AbaNovaEAbaRemovida()
    {
        var original = XlsxFixtures.Sales();
        var model = XlsxFixtures.Model(original);
        model.Sheets.Add(new SheetDto { Name = "Terceira", RowCount = 1000, ColumnCount = 26 });

        var (bytes, _) = XlsxWriter.Write(original, model);
        Assert.Equal(3, XlsxFixtures.Model(bytes).Sheets.Count);

        var shorter = XlsxFixtures.Model(bytes);
        shorter.Sheets.RemoveAt(2);
        var (fewer, _) = XlsxWriter.Write(bytes, shorter);
        Assert.Equal(2, XlsxFixtures.Model(fewer).Sheets.Count);
    }

    [Fact]
    public void PlanilhaNovaSemOriginal()
    {
        var model = new WorkbookDto();
        model.Sheets.Add(new SheetDto
        {
            Name = "Plan1",
            RowCount = 1000,
            ColumnCount = 26,
            Cells = { ["A1"] = new CellDto { Value = "Olá" }, ["B1"] = new CellDto { Value = 42d } },
        });

        var (bytes, report) = XlsxWriter.Write(null, model);

        Assert.Equal(2, report.CellsWritten);
        var sheet = XlsxFixtures.Sheet(XlsxFixtures.Model(bytes), "Plan1");
        Assert.Equal("Olá", sheet.Cells["A1"].Value);
        Assert.Equal(42d, sheet.Cells["B1"].Value);
    }

    [Fact]
    public void GravarSemAbaNenhumaEErro()
    {
        var problem = Assert.Throws<XlsxException>(() => XlsxWriter.Write(null, new WorkbookDto()));

        Assert.Contains("sem abas", problem.Message);
    }

    [Fact]
    public void IdaEVoltaPreservaOsValoresDeTodasAsCelulas()
    {
        var original = XlsxFixtures.Sales();
        var antes = XlsxFixtures.Model(original);

        var (bytes, _) = XlsxWriter.Write(original, antes);
        var depois = XlsxFixtures.Model(bytes);

        foreach (var sheet in antes.Sheets)
        {
            var other = XlsxFixtures.Sheet(depois, sheet.Name);
            Assert.Equal(sheet.Cells.Keys.OrderBy(key => key), other.Cells.Keys.OrderBy(key => key));
            foreach (var (reference, cell) in sheet.Cells)
            {
                Assert.True(cell.SameAs(other.Cells[reference]), $"{sheet.Name}!{reference} mudou");
            }
        }
    }
}

public class NumberFormatTests
{
    [Theory]
    [InlineData("\"R$\" #,##0.00", "currency", 2)]
    [InlineData("[$R$-416]#,##0.00", "currency", 2)]
    [InlineData("0.0%", "percent", 1)]
    // Data não tem casa decimal: a máscara não pede nenhuma, e `0` diria que o
    // usuário escolheu zero casas — coisa diferente de não haver escolha.
    [InlineData("dd/mm/yyyy", "date", null)]
    [InlineData("#,##0", "number", 0)]
    [InlineData("@", "text", null)]
    public void ReconheceAIntencaoDaMascara(string mask, string expected, int? decimals)
    {
        using var book = new XLWorkbook();
        var cell = book.Worksheets.Add("P").Cell("A1");
        cell.Style.NumberFormat.Format = mask;

        var (format, places) = NumberFormats.Read(cell.Style.NumberFormat);

        Assert.Equal(expected, format);
        if (decimals is not null) Assert.Equal(decimals, places);
    }

    [Fact]
    public void ColchetesNaoViramData()
    {
        // O `d` de `[$R$-416]` faria a coluna de preços virar datas de 1900.
        using var book = new XLWorkbook();
        var cell = book.Worksheets.Add("P").Cell("A1");
        cell.Style.NumberFormat.Format = "[$R$-416]#,##0.00";

        var (format, _) = NumberFormats.Read(cell.Style.NumberFormat);

        Assert.Equal("currency", format);
    }

    [Theory]
    [InlineData("General")]
    [InlineData("GENERAL")]
    [InlineData("Geral")]
    public void GeralPorExtensoNaoViraData(string mask)
    {
        // O LibreOffice grava a máscara `General` por extenso em vez de deixar o
        // código embutido zero. O "a" de Gener**a**l é o mesmo "a" de
        // `dd/mm/aaaa` — e sem tratamento a coluna de quantidades de toda
        // planilha vinda do LibreOffice abria cheia de datas de 1900.
        using var book = new XLWorkbook();
        var cell = book.Worksheets.Add("P").Cell("A1");
        cell.Style.NumberFormat.Format = mask;

        var (format, decimals) = NumberFormats.Read(cell.Style.NumberFormat);

        Assert.Null(format);
        Assert.Null(decimals);
    }

    [Fact]
    public void BarraInvertidaNaoViraData()
    {
        // `\d` é a letra d escapada, não o dia do mês.
        using var book = new XLWorkbook();
        var cell = book.Worksheets.Add("P").Cell("A1");
        cell.Style.NumberFormat.Format = "0.00\\ \"un\"";

        var (format, decimals) = NumberFormats.Read(cell.Style.NumberFormat);

        Assert.Equal("number", format);
        Assert.Equal(2, decimals);
    }

    [Fact]
    public void HoraNaoViraData()
    {
        // Um valor que é só hora mostrado como data daria 31/12/1899.
        using var book = new XLWorkbook();
        var cell = book.Worksheets.Add("P").Cell("A1");
        cell.Style.NumberFormat.NumberFormatId = 20;

        var (format, _) = NumberFormats.Read(cell.Style.NumberFormat);

        Assert.Null(format);
    }
}

public class UnitTests
{
    [Fact]
    public void LarguraPadraoDaOitentaQuatroPixels()
    {
        // 8,43 caracteres é a largura padrão de toda planilha, e ela vale 64
        // pixels. Errar a conta deixa toda coluna com a largura errada.
        Assert.Equal(64, Units.WidthToPixels(8.43));
    }

    [Fact]
    public void LarguraEIdaEVolta()
    {
        Assert.Equal(20, Math.Round(Units.PixelsToWidth(Units.WidthToPixels(20)), 2));
    }

    [Fact]
    public void PontosViramPixelsNaDensidadeDoCss()
    {
        Assert.Equal(96, Units.PointsToPixels(72));
    }
}

using ClosedXML.Excel;
using Librevia.Format.Docx;

namespace Librevia.Format.Xlsx;

/// <summary>
/// XLSX → modelo: valor, fórmula e alguns atributos de aparência; o resto vai ao
/// inventário. A fórmula sai como está no arquivo (nomes em inglês, vírgula), e o
/// analisador do lado TypeScript a traduz.
/// </summary>
public static class XlsxReader
{
    /// <summary>Uma coluna inteira formatada conta um milhão de células "usadas".</summary>
    private const int MaxCellsPerSheet = 200_000;

    public static XlsxOpenResult Read(byte[] bytes)
    {
        var inventory = new Inventory();

        using var stream = new MemoryStream(bytes, writable: false);
        XLWorkbook book;
        try
        {
            book = new XLWorkbook(stream);
        }
        catch (Exception problem) when (problem is not OutOfMemoryException)
        {
            throw new XlsxException(
                "Esta planilha não pôde ser lida: o arquivo parece corrompido ou não é uma planilha do Excel.");
        }

        using (book)
        {
            var workbook = new WorkbookDto();
            foreach (var sheet in book.Worksheets)
            {
                workbook.Sheets.Add(ReadSheet(sheet, inventory));
            }

            if (workbook.Sheets.Count == 0)
            {
                throw new XlsxException("Esta planilha não tem nenhuma aba.");
            }

            NoteBookWide(book, inventory);
            return new XlsxOpenResult(workbook, inventory);
        }
    }

    private static SheetDto ReadSheet(IXLWorksheet sheet, Inventory inventory)
    {
        var dto = new SheetDto
        {
            Name = sheet.Name,
            RowCount = 1000,
            ColumnCount = 26,
        };

        var used = sheet.LastCellUsed();
        if (used is not null)
        {
            // Uma tela de folga além do conteúdo.
            dto.RowCount = Math.Clamp(used.Address.RowNumber + 30, 1000, 1_000_000);
            dto.ColumnCount = Math.Clamp(used.Address.ColumnNumber + 5, 26, 16_384);
        }

        var read = 0;
        foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.AllContents | XLCellsUsedOptions.NormalFormats))
        {
            if (++read > MaxCellsPerSheet)
            {
                inventory.NoteInvisible(
                    $"a aba \"{sheet.Name}\" tem mais células do que o aplicativo abre; o excedente não foi carregado");
                break;
            }

            var converted = ReadCell(cell);
            if (converted is not null)
            {
                dto.Cells[cell.Address.ToStringRelative()] = converted;
            }
        }

        ReadDimensions(sheet, dto);
        ReadFrozen(sheet, dto);
        NoteSheetWide(sheet, inventory);
        return dto;
    }

    private static CellDto? ReadCell(IXLCell cell)
    {
        var dto = new CellDto();

        if (cell.HasFormula)
        {
            // O ClosedXML entrega a fórmula sem o `=`.
            dto.Formula = "=" + cell.FormulaA1;
        }

        dto.Value = ReadValue(cell);
        dto.Style = ReadStyle(cell);

        return dto.Value is null && dto.Formula is null && (dto.Style is null || dto.Style.IsEmpty) ? null : dto;
    }

    private static object? ReadValue(IXLCell cell) => cell.Value switch
    {
        { IsBlank: true } => null,
        { IsBoolean: true } value => value.GetBoolean(),
        { IsNumber: true } value => value.GetNumber(),
        { IsDateTime: true } value => value.GetDateTime().ToOADate(),
        { IsTimeSpan: true } value => value.GetTimeSpan().TotalDays,
        // Como texto, o mesmo formato do motor de fórmulas para `#DIV/0!` e companhia.
        { IsError: true } value => value.GetError().ToString(),
        { IsText: true } value => value.GetText(),
        _ => null,
    };

    private static CellStyleDto? ReadStyle(IXLCell cell)
    {
        var style = cell.Style;
        var dto = new CellStyleDto();

        if (style.Font.Bold) dto.Bold = true;
        if (style.Font.Italic) dto.Italic = true;
        if (style.Font.Underline != XLFontUnderlineValues.None) dto.Underline = true;

        dto.Color = HexOf(style.Font.FontColor, skipBlack: true);
        dto.Background = style.Fill.PatternType == XLFillPatternValues.None
            ? null
            : HexOf(style.Fill.BackgroundColor, skipBlack: false);

        dto.Align = style.Alignment.Horizontal switch
        {
            XLAlignmentHorizontalValues.Left => "left",
            XLAlignmentHorizontalValues.Center => "center",
            XLAlignmentHorizontalValues.Right => "right",
            _ => null,
        };

        var (format, decimals) = NumberFormats.Read(style.NumberFormat);
        dto.Format = format;
        dto.Decimals = decimals;

        var borders = new List<string>();
        if (style.Border.TopBorder != XLBorderStyleValues.None) borders.Add("top");
        if (style.Border.RightBorder != XLBorderStyleValues.None) borders.Add("right");
        if (style.Border.BottomBorder != XLBorderStyleValues.None) borders.Add("bottom");
        if (style.Border.LeftBorder != XLBorderStyleValues.None) borders.Add("left");
        if (borders.Count > 0) dto.Borders = borders;

        return dto.IsEmpty ? null : dto;
    }

    /// <summary>
    /// Cor de tema e indexada dependem da paleta, que o modelo não leva: nada, e não
    /// preto. O preto da fonte também sai, por ser o implícito de toda célula.
    /// </summary>
    private static string? HexOf(XLColor color, bool skipBlack)
    {
        if (!color.HasValue || color.ColorType != XLColorType.Color) return null;

        var value = color.Color;
        if (value.A == 0) return null;
        if (skipBlack && value is { R: 0, G: 0, B: 0 }) return null;

        return $"#{value.R:x2}{value.G:x2}{value.B:x2}";
    }

    private static void ReadDimensions(IXLWorksheet sheet, SheetDto dto)
    {
        foreach (var column in sheet.ColumnsUsed())
        {
            if (!column.Width.Equals(sheet.ColumnWidth))
            {
                dto.ColumnWidths[column.ColumnNumber() - 1] = Units.WidthToPixels(column.Width);
            }
        }

        foreach (var row in sheet.RowsUsed())
        {
            if (!row.Height.Equals(sheet.RowHeight))
            {
                dto.RowHeights[row.RowNumber() - 1] = Units.PointsToPixels(row.Height);
            }
        }
    }

    private static void ReadFrozen(IXLWorksheet sheet, SheetDto dto)
    {
        var view = sheet.SheetView;
        dto.FrozenRows = Math.Clamp(view.SplitRow, 0, 100);
        dto.FrozenColumns = Math.Clamp(view.SplitColumn, 0, 100);
    }

    /// <summary>
    /// O que a aba tem e a tela não mostra. É invisível, e não perda: o ClosedXML
    /// grava de volta o que não mexemos.
    /// </summary>
    private static void NoteSheetWide(IXLWorksheet sheet, Inventory inventory)
    {
        if (sheet.MergedRanges.Count > 0)
        {
            inventory.NoteInvisible("células mescladas (aparecem separadas na tela)");
        }
        if (sheet.ConditionalFormats.Any())
        {
            inventory.NoteInvisible("formatação condicional");
        }
        if (sheet.DataValidations.Any())
        {
            inventory.NoteInvisible("listas e validação de dados");
        }
        if (sheet.Pictures.Count > 0)
        {
            inventory.NoteInvisible("imagens");
        }
        if (sheet.Tables.Any())
        {
            inventory.NoteInvisible("tabelas formatadas");
        }
        if (sheet.AutoFilter.IsEnabled)
        {
            inventory.NoteInvisible("filtros");
        }
        // `CellsUsed(Comments)` devolve também as células com conteúdo sem comentário.
        if (sheet.CellsUsed(XLCellsUsedOptions.Comments, cell => cell.HasComment).Any())
        {
            inventory.NoteInvisible("comentários nas células");
        }
        if (sheet.Visibility != XLWorksheetVisibility.Visible)
        {
            inventory.NoteInvisible($"a aba \"{sheet.Name}\" está oculta no arquivo e aparece aqui como as outras");
        }
    }

    private static void NoteBookWide(XLWorkbook book, Inventory inventory)
    {
        if (book.DefinedNames.Any())
        {
            // O motor não conhece nomes definidos.
            inventory.NoteInvisible("intervalos nomeados (fórmulas que os usam não são calculadas)");
        }
    }
}

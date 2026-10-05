using ClosedXML.Excel;
using Librevia.Format.Docx;

namespace Librevia.Format.Xlsx;

/// <summary>
/// XLSX → model: values, formulas and a few appearance attributes; the rest goes to the inventory.
/// A formula comes out as stored in the file (English names, commas), and the TypeScript parser
/// translates it.
/// </summary>
public static class XlsxReader
{
    /// <summary>A whole formatted column counts a million "used" cells.</summary>
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
            // One screen of room beyond the content.
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
            // ClosedXML delivers the formula without the `=`.
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
        // As text, the same format as the formula engine for `#DIV/0!` and friends.
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
    /// Theme and indexed colors depend on the palette, which the model does not carry: nothing, not
    /// black. The font's black also goes, since it is every cell's implicit color.
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
    /// What the sheet has and the screen does not show. It is invisible, not lost: ClosedXML writes
    /// back what we do not touch.
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
        // `CellsUsed(Comments)` also returns cells with content and no comment.
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
            // The engine does not know defined names.
            inventory.NoteInvisible("intervalos nomeados (fórmulas que os usam não são calculadas)");
        }
    }
}

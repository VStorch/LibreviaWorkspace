using ClosedXML.Excel;
using Librevia.Format.Docx;

namespace Librevia.Format.Xlsx;

/// <summary>
/// XLSX → modelo do aplicativo.
/// </summary>
/// <remarks>
/// A leitura é deliberadamente **parcial**: o modelo do aplicativo representa
/// valor, fórmula e um punhado de atributos de aparência, e nada mais. O que
/// fica de fora não é ignorado em silêncio — vai para o inventário, que separa
/// o que continua no arquivo e não aparece na tela do que se perde ao gravar.
///
/// As fórmulas saem daqui **como estão no arquivo**: nomes em inglês e vírgula
/// separando argumentos. Quem traduz é o lado TypeScript, que já tem o
/// analisador — uma segunda implementação em C# seria duas gramáticas para
/// manter em acordo.
/// </remarks>
public static class XlsxReader
{
    /// <summary>
    /// Teto de células lidas por planilha.
    /// </summary>
    /// <remarks>
    /// Uma coluna inteira formatada gera um milhão de células "usadas" no
    /// critério do formato. Sem teto, abrir esse arquivo pareceria travamento.
    /// </remarks>
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
            // Folga de uma tela além do conteúdo, para o usuário ter onde
            // digitar sem precisar inserir linha antes.
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
            // O `=` faz parte do que o aplicativo guarda; o ClosedXML entrega sem.
            dto.Formula = "=" + cell.FormulaA1;
        }

        dto.Value = ReadValue(cell);
        dto.Style = ReadStyle(cell);

        // Célula sem nada não entra no mapa esparso.
        return dto.Value is null && dto.Formula is null && (dto.Style is null || dto.Style.IsEmpty) ? null : dto;
    }

    private static object? ReadValue(IXLCell cell) => cell.Value switch
    {
        { IsBlank: true } => null,
        { IsBoolean: true } value => value.GetBoolean(),
        { IsNumber: true } value => value.GetNumber(),
        // Data vira número de série, que é como a planilha de fato a guarda.
        { IsDateTime: true } value => value.GetDateTime().ToOADate(),
        { IsTimeSpan: true } value => value.GetTimeSpan().TotalDays,
        // Erro guardado como texto: é o mesmo formato em que o motor de
        // fórmulas do aplicativo representa `#DIV/0!` e companhia.
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
    /// Cor em <c>#rrggbb</c>, ou nada.
    /// </summary>
    /// <remarks>
    /// Cor de tema e cor indexada dependem da paleta do arquivo, que o modelo
    /// do aplicativo não carrega. Devolver preto para elas pintaria de preto
    /// texto que estava colorido — pior que não pintar.
    ///
    /// Preto puro na fonte também sai fora: é a cor implícita de toda célula, e
    /// gravá-la em cada uma encheria o arquivo de estilo que não diz nada.
    /// </remarks>
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
    /// O que existe na aba e o aplicativo não mostra.
    /// </summary>
    /// <remarks>
    /// Tudo aqui é **invisibilidade**, e não perda: o ClosedXML grava de volta o
    /// que não mexemos. Medido: uma parte de XML personalizada injetada à mão
    /// sobrevive à ida e volta. O que ele regenera é a planilha em si, e é por
    /// isso que a lista abaixo existe.
    /// </remarks>
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
        // O predicado não é redundante: `CellsUsed(Comments)` devolve as células
        // com conteúdo mesmo quando nenhuma tem comentário, e sem ele toda
        // planilha comum sairia com um aviso falso.
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
            // Estes importam mais que os outros: uma fórmula que usa nome
            // definido não calcula, porque o motor não os conhece.
            inventory.NoteInvisible("intervalos nomeados (fórmulas que os usam não são calculadas)");
        }
    }
}

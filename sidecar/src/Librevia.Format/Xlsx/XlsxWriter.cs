using ClosedXML.Excel;

namespace Librevia.Format.Xlsx;

/// <summary>
/// Modelo → XLSX, escrevendo só o que mudou. O ClosedXML preserva as partes que não
/// modela, mas regenera a planilha: a gravação relê o original e toca só as células
/// que mudaram, para não apagar o que o modelo não representa.
/// </summary>
public static class XlsxWriter
{
    public sealed record Result(int Sheets, int CellsWritten, int CellsCleared, int CellsPreserved);

    public static (byte[] Bytes, Result Report) Write(byte[]? original, WorkbookDto model)
    {
        if (model.Sheets.Count == 0)
        {
            throw new XlsxException("Não há nada para gravar: a planilha ficou sem abas.");
        }

        // O ClosedXML lê as partes que faltam no `SaveAs`: fechado antes, "Cannot access a closed Stream".
        using var source = new MemoryStream(original ?? [], writable: false);
        using var book = Open(source, original is not null && original.Length > 0);
        var before = original is null ? null : XlsxReader.Read(original).Workbook;

        var written = 0;
        var cleared = 0;
        var preserved = 0;

        SyncSheets(book, model);

        for (var index = 0; index < model.Sheets.Count; index++)
        {
            var wanted = model.Sheets[index];
            var sheet = book.Worksheet(index + 1);
            var previous = before?.Sheets.ElementAtOrDefault(index);

            // Pela posição, e não pelo nome: aba renomeada não é aba nova.
            var (w, c, p) = SyncCells(sheet, wanted, previous);
            written += w;
            cleared += c;
            preserved += p;

            SyncDimensions(sheet, wanted, previous);
            SyncFrozen(sheet, wanted, previous);
        }

        book.CalculateMode = XLCalculateMode.Auto;

        using var output = new MemoryStream();
        book.SaveAs(output);
        return (output.ToArray(), new Result(model.Sheets.Count, written, cleared, preserved));
    }

    private static XLWorkbook Open(Stream source, bool hasOriginal)
    {
        if (!hasOriginal) return new XLWorkbook();

        try
        {
            return new XLWorkbook(source);
        }
        catch (Exception problem) when (problem is not OutOfMemoryException)
        {
            throw new XlsxException(
                "O arquivo original não pôde ser lido para gravar por cima. Use \"Salvar como\" para gravar num arquivo novo.");
        }
    }

    private static void SyncSheets(XLWorkbook book, WorkbookDto model)
    {
        while (book.Worksheets.Count > model.Sheets.Count)
        {
            book.Worksheet(book.Worksheets.Count).Delete();
        }

        while (book.Worksheets.Count < model.Sheets.Count)
        {
            // Provisório: o nome final pode ainda ser o de outra aba.
            book.Worksheets.Add($"__nova{book.Worksheets.Count + 1}");
        }

        // Em duas passadas: o ClosedXML recusa o nome que outra aba ainda tem.
        for (var index = 0; index < model.Sheets.Count; index++)
        {
            var sheet = book.Worksheet(index + 1);
            if (sheet.Name != model.Sheets[index].Name) sheet.Name = $"__temp{index}__";
        }

        for (var index = 0; index < model.Sheets.Count; index++)
        {
            var sheet = book.Worksheet(index + 1);
            if (sheet.Name != model.Sheets[index].Name) sheet.Name = model.Sheets[index].Name;
        }
    }

    private static (int Written, int Cleared, int Preserved) SyncCells(
        IXLWorksheet sheet,
        SheetDto wanted,
        SheetDto? previous)
    {
        var written = 0;
        var preserved = 0;

        foreach (var (reference, cell) in wanted.Cells)
        {
            var before = previous?.Cells.GetValueOrDefault(reference);
            if (cell.SameAs(before))
            {
                preserved++;
                continue;
            }

            WriteCell(sheet.Cell(reference), cell, before);
            written++;
        }

        var cleared = 0;
        if (previous is not null)
        {
            foreach (var reference in previous.Cells.Keys)
            {
                if (wanted.Cells.ContainsKey(reference)) continue;

                // Conteúdo e formato, e não a célula: o da linha e o da coluna ficam.
                sheet.Cell(reference).Clear(XLClearOptions.Contents | XLClearOptions.NormalFormats);
                cleared++;
            }
        }

        return (written, cleared, preserved);
    }

    private static void WriteCell(IXLCell target, CellDto cell, CellDto? before)
    {
        if (cell.Formula is not null)
        {
            var formula = cell.Formula.StartsWith('=') ? cell.Formula[1..] : cell.Formula;
            if (before?.Formula != cell.Formula) target.FormulaA1 = formula;
        }
        else
        {
            if (before?.Formula is not null) target.FormulaA1 = string.Empty;
            WriteValue(target, cell.Value);
        }

        WriteStyle(target, cell.Style, before?.Style);
    }

    private static void WriteValue(IXLCell target, object? value)
    {
        switch (value)
        {
            case null:
                target.Clear(XLClearOptions.Contents);
                break;
            case bool flag:
                target.Value = flag;
                break;
            case string text:
                // Senão o Excel interpreta o `=` ao reabrir.
                target.SetValue(text);
                break;
            default:
                target.Value = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                break;
        }
    }

    /// <summary>Só os atributos que mudaram: o estilo inteiro apagaria fonte, recuo e quebra.</summary>
    private static void WriteStyle(IXLCell target, CellStyleDto? style, CellStyleDto? before)
    {
        var wanted = style ?? new CellStyleDto();
        var had = before ?? new CellStyleDto();
        if (CellStyleDto.Same(wanted, had)) return;

        if (wanted.Bold != had.Bold) target.Style.Font.Bold = wanted.Bold == true;
        if (wanted.Italic != had.Italic) target.Style.Font.Italic = wanted.Italic == true;
        if (wanted.Underline != had.Underline)
        {
            target.Style.Font.Underline =
                wanted.Underline == true ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
        }

        if (wanted.Color != had.Color)
        {
            target.Style.Font.FontColor = wanted.Color is null ? XLColor.Black : ColorOf(wanted.Color);
        }

        if (wanted.Background != had.Background)
        {
            if (wanted.Background is null) target.Style.Fill.PatternType = XLFillPatternValues.None;
            else target.Style.Fill.SetBackgroundColor(ColorOf(wanted.Background));
        }

        if (wanted.Align != had.Align)
        {
            target.Style.Alignment.Horizontal = wanted.Align switch
            {
                "left" => XLAlignmentHorizontalValues.Left,
                "center" => XLAlignmentHorizontalValues.Center,
                "right" => XLAlignmentHorizontalValues.Right,
                _ => XLAlignmentHorizontalValues.General,
            };
        }

        if (!NumberFormats.Matches(target.Style.NumberFormat, wanted.Format, wanted.Decimals))
        {
            var mask = NumberFormats.Mask(wanted.Format, wanted.Decimals);
            if (mask is null) target.Style.NumberFormat.NumberFormatId = 0;
            else target.Style.NumberFormat.Format = mask;
        }

        if (!SameBorders(wanted.Borders, had.Borders))
        {
            var sides = wanted.Borders ?? [];
            target.Style.Border.TopBorder = Side(sides.Contains("top"));
            target.Style.Border.RightBorder = Side(sides.Contains("right"));
            target.Style.Border.BottomBorder = Side(sides.Contains("bottom"));
            target.Style.Border.LeftBorder = Side(sides.Contains("left"));
        }
    }

    private static XLBorderStyleValues Side(bool present) =>
        present ? XLBorderStyleValues.Thin : XLBorderStyleValues.None;

    private static bool SameBorders(List<string>? left, List<string>? right)
    {
        var a = left ?? [];
        var b = right ?? [];
        return a.Count == b.Count && !a.Except(b, StringComparer.Ordinal).Any();
    }

    private static XLColor ColorOf(string hex)
    {
        var text = hex.TrimStart('#');
        if (text.Length != 6 || !int.TryParse(text, System.Globalization.NumberStyles.HexNumber, null, out _))
        {
            return XLColor.Black;
        }

        return XLColor.FromHtml("#" + text);
    }

    private static void SyncDimensions(IXLWorksheet sheet, SheetDto wanted, SheetDto? previous)
    {
        foreach (var (index, pixels) in wanted.ColumnWidths)
        {
            if (previous is not null
                && previous.ColumnWidths.TryGetValue(index, out var was)
                && Math.Abs(was - pixels) < 0.5)
            {
                continue;
            }

            sheet.Column(index + 1).Width = Units.PixelsToWidth(pixels);
        }

        foreach (var (index, pixels) in wanted.RowHeights)
        {
            if (previous is not null
                && previous.RowHeights.TryGetValue(index, out var was)
                && Math.Abs(was - pixels) < 0.5)
            {
                continue;
            }

            sheet.Row(index + 1).Height = Units.PixelsToPoints(pixels);
        }
    }

    private static void SyncFrozen(IXLWorksheet sheet, SheetDto wanted, SheetDto? previous)
    {
        if (previous is not null
            && previous.FrozenRows == wanted.FrozenRows
            && previous.FrozenColumns == wanted.FrozenColumns)
        {
            return;
        }

        sheet.SheetView.Freeze(wanted.FrozenRows, wanted.FrozenColumns);
    }
}

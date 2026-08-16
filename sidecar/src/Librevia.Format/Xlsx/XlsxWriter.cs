using ClosedXML.Excel;

namespace Librevia.Format.Xlsx;

/// <summary>
/// Modelo do aplicativo → XLSX, escrevendo **só o que mudou**.
/// </summary>
/// <remarks>
/// O princípio é o mesmo da Fase 4: o que o usuário não editou volta como
/// estava. A execução é diferente, e a diferença vem de uma medição.
///
/// No DOCX, a preservação é feita à mão, parte por parte, porque a biblioteca
/// regenerava tudo. No XLSX, o ClosedXML **preserva as partes que não modela** —
/// uma parte de XML personalizada injetada à mão sobrevive à ida e volta. O que
/// ele regenera é a planilha em si.
///
/// Então o risco aqui não é o pacote: é a célula. Escrever todas as células a
/// cada gravação apagaria fonte, tamanho, alinhamento vertical, recuo e
/// bordas diagonais — tudo que o modelo do aplicativo não representa. Por isso a
/// gravação relê o arquivo original com o mesmo leitor, compara célula a célula,
/// e só toca no que de fato mudou.
/// </remarks>
public static class XlsxWriter
{
    public sealed record Result(int Sheets, int CellsWritten, int CellsCleared, int CellsPreserved);

    public static (byte[] Bytes, Result Report) Write(byte[]? original, WorkbookDto model)
    {
        if (model.Sheets.Count == 0)
        {
            throw new XlsxException("Não há nada para gravar: a planilha ficou sem abas.");
        }

        // O fluxo de origem precisa viver até o `SaveAs`: o ClosedXML guarda
        // referência a ele e só lê as partes que faltam na hora de gravar.
        // Fechá-lo antes derruba a gravação com "Cannot access a closed Stream".
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

            // A aba comparada é a de mesma **posição**, não a de mesmo nome:
            // renomear uma aba não pode fazer o conteúdo dela parecer novo.
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

    /// <summary>Acerta quantidade, ordem e nome das abas.</summary>
    private static void SyncSheets(XLWorkbook book, WorkbookDto model)
    {
        while (book.Worksheets.Count > model.Sheets.Count)
        {
            book.Worksheet(book.Worksheets.Count).Delete();
        }

        while (book.Worksheets.Count < model.Sheets.Count)
        {
            // Nome provisório: o acerto de nomes vem logo abaixo, e usar o nome
            // final aqui esbarraria numa aba que ainda não foi renomeada.
            book.Worksheets.Add($"__nova{book.Worksheets.Count + 1}");
        }

        // Renomear em duas passadas: dar a uma aba o nome que outra ainda tem é
        // um erro do ClosedXML, e trocar duas abas de nome é um caso normal.
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

                // Some do modelo quer dizer apagada pelo usuário. Limpar
                // conteúdo e formato, e não a célula inteira, evita mexer no que
                // é da linha ou da coluna.
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
                // Texto que começa com `=` sem ser fórmula precisa ir como
                // texto, senão o Excel o interpreta ao reabrir.
                target.SetValue(text);
                break;
            default:
                target.Value = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                break;
        }
    }

    /// <summary>
    /// Aplica só os atributos que mudaram.
    /// </summary>
    /// <remarks>
    /// Escrever o estilo inteiro apagaria o que o modelo não carrega: nome e
    /// tamanho da fonte, alinhamento vertical, recuo, quebra de texto. O usuário
    /// perderia a aparência de uma célula por ter mudado a cor dela.
    /// </remarks>
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

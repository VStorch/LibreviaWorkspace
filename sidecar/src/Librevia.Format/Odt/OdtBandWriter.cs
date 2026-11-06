using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A header or footer from another format (<c>band.ts</c>) → the content of
/// <c>style:header</c>/<c>style:footer</c>. The grid becomes a table; the three columns, a paragraph
/// with center and right tab stops when each has one line, or aligned paragraphs when they have
/// more. The anchored objects go in the first paragraph, positioned on the sheet.
/// </summary>
internal sealed class OdtBandWriter(OdtWriteContext context, OdtBlockWriter blocks)
{
    private const string CellPadding = "1.9mm";
    private const string Standard = OdtStyles.DefaultParagraphStyle;
    private const string FloatsOnlyFontSize = "1pt";
    private const double MinColumnMm = 1;
    private const double DefaultImagePixels = 1;

    private static readonly (string Side, char Letter)[] Sides = [("top", 't'), ("bottom", 'b'), ("left", 'l'), ("right", 'r')];

    public List<XElement> Write(BandDto band, double widthMm, bool header)
    {
        var parts = new List<XElement>();
        if (band.Rows is { Count: > 0 } rows) parts.Add(Grid(rows, widthMm));
        parts.AddRange(Columns(band, widthMm, header && band.Rule));

        context.InBand = true;
        var frames = new OdtFloatWriter(context, blocks).Write(band.Floats ?? []).ToList();
        context.InBand = false;
        if (frames.Count > 0) PlaceFrames(parts, frames);
        return parts.Count == 0 ? [new XElement(Ns.Text + "p")] : parts;
    }

    private IEnumerable<XElement> Columns(BandDto band, double widthMm, bool rule)
    {
        var columns = new[] { (band.Left, "left"), (band.Center, "center"), (band.Right, "right") }
            .Where(column => column.Item1.Count > 0).ToList();
        if (columns.Count == 0) return [];
        if (columns.All(column => LinesOf(column.Item1).Count == 1)) return [Tabbed(band, widthMm, rule)];

        var paragraphs = new List<XElement>();
        for (var index = 0; index < columns.Count; index++)
        {
            var lines = LinesOf(columns[index].Item1);
            for (var line = 0; line < lines.Count; line++)
            {
                var last = index == columns.Count - 1 && line == lines.Count - 1;
                paragraphs.Add(Paragraph(Properties(columns[index].Item2, last && rule), Pieces(lines[line])));
            }
        }

        return paragraphs;
    }

    /// <summary>Left, a tab, center, a tab, right; without trailing tabs that carry nothing.</summary>
    private XElement Tabbed(BandDto band, double widthMm, bool rule)
    {
        var properties = Properties(null, rule);
        properties.Add(new XElement(
            Ns.Style + "tab-stops",
            TabStop("center", widthMm / 2),
            TabStop("right", widthMm)));
        var content = new List<List<XNode>> { Pieces(band.Left), Pieces(band.Center), Pieces(band.Right) };
        while (content.Count > 1 && content[^1].Count == 0) content.RemoveAt(content.Count - 1);

        var nodes = new List<XNode>();
        for (var index = 0; index < content.Count; index++)
        {
            if (index > 0) nodes.Add(new XElement(Ns.Text + "tab"));
            nodes.AddRange(content[index]);
        }

        return Paragraph(properties, nodes);
    }

    private static XElement TabStop(string type, double positionMm) => new(
        Ns.Style + "tab-stop",
        new XAttribute(Ns.Style + "type", type),
        new XAttribute(Ns.Style + "position", OdtFormatWriter.Mm(positionMm)!));

    private static XElement Properties(string? align, bool rule)
    {
        var properties = OdtFormatWriter.Paragraph(new ParagraphFormat(TextAlign: align, SpaceBefore: 0, SpaceAfter: 0));
        if (rule) properties.SetAttributeValue(Ns.Fo + "border-bottom", OdtStyles.DefaultBorder);
        return properties;
    }

    private XElement Paragraph(XElement properties, IEnumerable<XNode> content, XElement? text = null) => new(
        Ns.Text + "p",
        new XAttribute(Ns.Text + "style-name", context.Automatic.Style(OdtAutomaticStyles.Band, OdtStyles.ParagraphFamily, Standard, properties, text)),
        content);

    /// <summary>Into the first paragraph outside the grid, or a paragraph of their own that takes no room.</summary>
    private void PlaceFrames(List<XElement> parts, List<XElement> frames)
    {
        if (parts.FirstOrDefault(part => part.Name == Ns.Text + "p") is { } first)
        {
            first.AddFirst(frames);
            return;
        }

        var properties = OdtFormatWriter.Paragraph(new ParagraphFormat(
            SpaceBefore: 0, SpaceAfter: 0, LineSpacing: new LineSpacingDto("multiple", Factor: 1)));
        var tiny = new XElement(Ns.Style + "text-properties", new XAttribute(Ns.Fo + "font-size", FloatsOnlyFontSize));
        parts.Add(Paragraph(properties, frames, tiny));
    }

    /// <summary>Each line of a column starts at a piece marked <c>line</c>.</summary>
    private static List<List<PieceDto>> LinesOf(List<PieceDto> pieces)
    {
        var lines = new List<List<PieceDto>>();
        foreach (var piece in pieces)
        {
            if (piece.Line || lines.Count == 0) lines.Add([]);
            lines[^1].Add(piece);
        }

        return lines;
    }

    private List<XNode> Pieces(List<PieceDto> pieces) => [.. pieces.SelectMany(Piece)];

    private IEnumerable<XNode> Piece(PieceDto piece)
    {
        List<XNode> inner = piece.Kind switch
        {
            "pageNumber" => [OdtFields.Create("page-number", "1")],
            "totalPages" => [OdtFields.Create("page-count", "1")],
            "image" => ImageOf(piece) is { } frame ? [frame] : [],
            _ => OdtInlineWriter.PlainText(piece.Text ?? string.Empty, pageFields: !piece.Literal),
        };
        if (piece.Kind == "image" || inner.Count == 0) return inner;

        var format = new CharacterFormat(
            FontFamily: piece.FontFamily,
            FontSizePt: Attr.Points(piece.FontSize),
            Bold: piece.Bold ? true : null,
            Italic: piece.Italic ? true : null,
            Color: CharacterFormat.ColorOf(piece.Color));
        var properties = OdtFormatWriter.Text(format, context.Styles.Fonts);
        if (!properties.HasAttributes) return inner;
        var style = context.Automatic.Style(OdtAutomaticStyles.Text, OdtStyles.TextFamily, null, properties);
        return [new XElement(Ns.Text + "span", new XAttribute(Ns.Text + "style-name", style), inner)];
    }

    private XElement? ImageOf(PieceDto piece)
    {
        if (OdtInlineWriter.Picture(context, piece.Src) is not { } path) return null;
        var natural = OdtInlineWriter.NaturalSize(piece.Src);
        return OdtInlineWriter.CharacterFrame(
            context, path, piece.Width ?? natural?.Width ?? DefaultImagePixels, piece.Height ?? natural?.Height ?? DefaultImagePixels);
    }

    private XElement Grid(List<BandRowDto> rows, double widthMm)
    {
        var grid = BandGrid.Of(rows);
        var widths = grid.Widths(widthMm, MinColumnMm);
        var table = new XElement(
            Ns.Table + "table",
            new XAttribute(Ns.Table + "name", context.NewTableName()),
            new XAttribute(Ns.Table + "style-name", context.Automatic.Style(
                OdtAutomaticStyles.Table, "table", null,
                new XElement(
                    Ns.Style + "table-properties",
                    new XAttribute(Ns.Style + "width", OdtFormatWriter.Mm(widths.Sum())!),
                    new XAttribute(Ns.Table + "align", "left")))));
        foreach (var width in widths)
        {
            var column = new XElement(Ns.Style + "table-column-properties", new XAttribute(Ns.Style + "column-width", OdtFormatWriter.Mm(width)!));
            table.Add(new XElement(Ns.Table + "table-column", new XAttribute(Ns.Table + "style-name", context.Automatic.Style(OdtAutomaticStyles.Column, "table-column", null, column))));
        }

        foreach (var slots in grid.Rows)
        {
            table.Add(new XElement(Ns.Table + "table-row", slots.Select(slot => slot is null ? new XElement(Ns.Table + "covered-table-cell") : Cell(slot))));
        }

        return table;
    }

    private XElement Cell(BandGrid.Slot slot)
    {
        var properties = new XElement(
            Ns.Style + "table-cell-properties",
            new XAttribute(Ns.Fo + "padding-left", CellPadding),
            new XAttribute(Ns.Fo + "padding-right", CellPadding),
            new XAttribute(Ns.Fo + "padding-top", "0mm"),
            new XAttribute(Ns.Fo + "padding-bottom", "0mm"));
        foreach (var (side, letter) in Sides)
        {
            properties.Add(new XAttribute(Ns.Fo + "border-" + side, slot.Cell.Borders.Contains(letter) ? OdtStyles.DefaultBorder : "none"));
        }

        var cell = new XElement(
            Ns.Table + "table-cell",
            new XAttribute(Ns.Table + "style-name", context.Automatic.Style(OdtAutomaticStyles.Cell, "table-cell", null, properties)),
            new XAttribute(Ns.Office + "value-type", "string"));
        if (slot.Colspan > 1) cell.Add(new XAttribute(Ns.Table + "number-columns-spanned", slot.Colspan));
        if (slot.Rowspan > 1) cell.Add(new XAttribute(Ns.Table + "number-rows-spanned", slot.Rowspan));

        var lines = LinesOf(slot.Cell.Pieces);
        if (lines.Count == 0) lines.Add([]);
        foreach (var line in lines) cell.Add(Paragraph(Properties(slot.Cell.Align, rule: false), Pieces(line)));
        return cell;
    }
}

/// <summary>
/// A band's grid with one slot per column in every row, as ODF wants it. A cell states its fraction
/// of the width; a single-column cell gives the column, and a merged one shares what is left among
/// the columns nobody measured.
/// </summary>
internal sealed class BandGrid
{
    public sealed record Slot(BandCellDto Cell, int Colspan, int Rowspan);

    private BandGrid(int columns, List<List<Slot?>> rows)
    {
        Columns = columns;
        Rows = rows;
    }

    public int Columns { get; }

    public List<List<Slot?>> Rows { get; }

    public static BandGrid Of(List<BandRowDto> rows)
    {
        var below = new List<int>();
        var grid = new List<List<Slot?>>();
        foreach (var row in rows)
        {
            var slots = new List<Slot?>();
            foreach (var cell in row.Cells)
            {
                Covered(slots, below);
                var slot = new Slot(cell, Math.Max(1, cell.Span), Math.Max(1, cell.RowSpan));
                for (var column = 0; column < slot.Colspan; column++)
                {
                    while (below.Count <= slots.Count) below.Add(0);
                    below[slots.Count] = slot.Rowspan - 1;
                    slots.Add(column == 0 ? slot : null);
                }
            }

            Covered(slots, below);
            grid.Add(slots);
        }

        var columns = grid.Count == 0 ? 0 : grid.Max(slots => slots.Count);
        foreach (var slots in grid)
        {
            while (slots.Count < columns) slots.Add(null);
        }

        return new BandGrid(columns, grid);
    }

    private static void Covered(List<Slot?> slots, List<int> below)
    {
        while (slots.Count < below.Count && below[slots.Count] > 0)
        {
            below[slots.Count]--;
            slots.Add(null);
        }
    }

    public List<double> Widths(double widthMm, double minimumMm)
    {
        var known = new double?[Columns];
        var spans = new List<(int Column, int Span, double Width)>();
        foreach (var slots in Rows)
        {
            for (var column = 0; column < slots.Count; column++)
            {
                if (slots[column] is not { } slot || slot.Cell.Width <= 0) continue;
                var width = slot.Cell.Width * widthMm;
                if (slot.Colspan == 1 && known[column] is null) known[column] = width;
                else spans.Add((column, slot.Colspan, width));
            }
        }

        foreach (var (column, span, width) in spans) ShareSpan(known, column, span, width, minimumMm);
        var measured = known.Sum(value => value ?? 0);
        var unknown = known.Count(value => value is null);
        var share = unknown == 0 ? 0 : Math.Max(minimumMm, (widthMm - measured) / unknown);
        return [.. known.Select(value => value ?? share)];
    }

    private static void ShareSpan(double?[] known, int column, int span, double width, double minimumMm)
    {
        var end = Math.Min(known.Length, column + span);
        var missing = 0;
        var taken = 0.0;
        for (var index = column; index < end; index++)
        {
            if (known[index] is { } value) taken += value;
            else missing++;
        }

        if (missing == 0) return;
        for (var index = column; index < end; index++) known[index] ??= Math.Max(minimumMm, (width - taken) / missing);
    }
}

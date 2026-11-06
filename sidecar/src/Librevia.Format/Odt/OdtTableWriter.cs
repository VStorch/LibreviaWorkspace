using System.Globalization;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inverse of <see cref="OdtTableReader"/>: a table → <c>table:table</c> with one slot per column
/// in every row, <c>table:covered-table-cell</c> where a span reaches, the leading header rows in
/// <c>table:table-header-rows</c>, and the shading and borders in the cell's automatic style.
/// </summary>
internal sealed class OdtTableWriter(OdtWriteContext context, OdtBlockWriter blocks)
{
    /// <summary>The thin line the editor draws on a cell that states none; see <c>export-odt.ts</c>.</summary>
    private const string DefaultBorder = "0.5pt solid #000000";

    private const string CellPadding = "0.97mm";

    private static readonly string[] Sides = ["top", "right", "bottom", "left"];

    private sealed record Slot(Node Cell, int Colspan, int Rowspan);

    public XElement Write(Node table, bool breakBefore)
    {
        var rows = (table.Content ?? []).Where(row => row.Type == "tableRow").ToList();
        var (columns, grid) = Grid(rows);
        var widths = ColumnWidths(rows, columns);

        var element = new XElement(
            Ns.Table + "table",
            new XAttribute(Ns.Table + "name", context.NewTableName()),
            new XAttribute(Ns.Table + "style-name", TableStyle(widths, breakBefore)));
        element.Add(Columns(widths, columns));

        var header = new XElement(Ns.Table + "table-header-rows");
        for (var index = 0; index < rows.Count; index++)
        {
            var row = Row(grid[index]);
            var isHeader = index == header.Elements().Count() && rows[index].Content?.All(cell => cell.Type == "tableHeader") == true;
            if (isHeader) header.Add(row);
            else element.Add(row);
        }

        if (header.HasElements) element.Elements(Ns.Table + "table-column").Last().AddAfterSelf(header);
        return element;
    }

    /// <summary>
    /// One slot per column: a cell, or <c>null</c> where a cell above or to the left spans. Every row
    /// gets the same count.
    /// </summary>
    private static (int Columns, List<List<Slot?>> Grid) Grid(List<Node> rows)
    {
        var below = new List<int>();
        var grid = new List<List<Slot?>>();
        foreach (var row in rows)
        {
            var slots = new List<Slot?>();
            void Covered()
            {
                while (slots.Count < below.Count && below[slots.Count] > 0)
                {
                    below[slots.Count]--;
                    slots.Add(null);
                }
            }

            foreach (var cell in row.Content ?? [])
            {
                Covered();
                var slot = new Slot(cell, Math.Max(1, Attr.Int(cell, "colspan") ?? 1), Math.Max(1, Attr.Int(cell, "rowspan") ?? 1));
                for (var column = 0; column < slot.Colspan; column++)
                {
                    var at = slots.Count;
                    while (below.Count <= at) below.Add(0);
                    below[at] = slot.Rowspan - 1;
                    slots.Add(column == 0 ? slot : null);
                }
            }

            Covered();
            grid.Add(slots);
        }

        var columns = grid.Count == 0 ? 0 : grid.Max(slots => slots.Count);
        foreach (var slots in grid)
        {
            while (slots.Count < columns) slots.Add(null);
        }

        return (Math.Max(1, columns), grid);
    }

    /// <summary>From the first row, in pixels, when every cell measures its columns.</summary>
    private static List<double>? ColumnWidths(List<Node> rows, int columns)
    {
        var widths = new List<double>();
        foreach (var cell in rows.FirstOrDefault()?.Content ?? [])
        {
            if (Attr.Node(cell, "colwidth") is not JsonArray declared || declared.Count == 0) return null;
            foreach (var width in declared)
            {
                if (width is not JsonValue value ||
                    !double.TryParse(value.ToJsonString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var pixels) ||
                    pixels <= 0)
                {
                    return null;
                }

                widths.Add(pixels);
            }
        }

        return widths.Count == columns ? widths : null;
    }

    private string TableStyle(List<double>? widths, bool breakBefore)
    {
        var properties = new XElement(Ns.Style + "table-properties");
        if (widths is null)
        {
            properties.SetAttributeValue(Ns.Table + "align", "margins");
        }
        else
        {
            properties.SetAttributeValue(Ns.Style + "width", OdtFormatWriter.PixelsToMm(widths.Sum()));
            properties.SetAttributeValue(Ns.Table + "align", "left");
        }

        if (breakBefore) properties.SetAttributeValue(Ns.Fo + "break-before", ParagraphFormat.Page);
        return context.Automatic.Style(OdtAutomaticStyles.Table, "table", null, properties);
    }

    private IEnumerable<XElement> Columns(List<double>? widths, int columns)
    {
        if (widths is null)
        {
            yield return new XElement(Ns.Table + "table-column", new XAttribute(Ns.Table + "number-columns-repeated", columns));
            yield break;
        }

        foreach (var width in widths)
        {
            var properties = new XElement(Ns.Style + "table-column-properties", new XAttribute(Ns.Style + "column-width", OdtFormatWriter.PixelsToMm(width)));
            var style = context.Automatic.Style(OdtAutomaticStyles.Column, "table-column", null, properties);
            yield return new XElement(Ns.Table + "table-column", new XAttribute(Ns.Table + "style-name", style));
        }
    }

    private XElement Row(List<Slot?> slots)
    {
        var row = new XElement(Ns.Table + "table-row");
        foreach (var slot in slots) row.Add(slot is null ? new XElement(Ns.Table + "covered-table-cell") : Cell(slot));
        return row;
    }

    private XElement Cell(Slot slot)
    {
        var cell = new XElement(
            Ns.Table + "table-cell",
            new XAttribute(Ns.Table + "style-name", CellStyle(slot.Cell)),
            new XAttribute(Ns.Office + "value-type", "string"));
        if (slot.Colspan > 1) cell.Add(new XAttribute(Ns.Table + "number-columns-spanned", slot.Colspan));
        if (slot.Rowspan > 1) cell.Add(new XAttribute(Ns.Table + "number-rows-spanned", slot.Rowspan));

        foreach (var child in slot.Cell.Content ?? []) cell.Add(blocks.Write(child, flat: true));
        if (!cell.HasElements) cell.Add(new XElement(Ns.Text + "p"));
        return cell;
    }

    private string CellStyle(Node cell)
    {
        var properties = new XElement(Ns.Style + "table-cell-properties", new XAttribute(Ns.Fo + "padding", context.KeptCellPadding ?? CellPadding));
        if (CharacterFormat.ColorOf(Attr.String(cell, "shading")) is { } shading)
        {
            properties.SetAttributeValue(Ns.Fo + "background-color", shading);
        }

        var stated = Borders(Attr.String(cell, "borders"));
        foreach (var side in Sides)
        {
            properties.SetAttributeValue(Ns.Fo + "border-" + side, stated.GetValueOrDefault(side, DefaultBorder));
        }

        return context.Automatic.Style(OdtAutomaticStyles.Cell, "table-cell", null, properties);
    }

    /// <summary><c>top:double,1.5,#ff0000;…</c> → <c>1.5pt double #ff0000</c> per side.</summary>
    private static Dictionary<string, string> Borders(string? text)
    {
        var sides = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in (text ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var colon = entry.IndexOf(':');
            if (colon <= 0) continue;
            var parts = entry[(colon + 1)..].Split(',');
            var kind = parts[0].Trim();
            if (kind == "none")
            {
                sides[entry[..colon].Trim()] = "none";
                continue;
            }

            var width = parts.Length > 1 ? parts[1].Trim() : "0.5";
            var color = parts.Length > 2 ? parts[2].Trim() : "#000000";
            var style = kind is "double" or "dotted" or "dashed" ? kind : "solid";
            sides[entry[..colon].Trim()] = $"{width}pt {style} {color}";
        }

        return sides;
    }
}

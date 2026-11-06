using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>table:table</c> → <c>table</c>, with the attributes <see cref="BodyReader"/> gives a DOCX
/// table: <c>colspan</c> and <c>rowspan</c> on every cell, the column widths in CSS pixels, the
/// shading and the borders in <c>table-format.ts</c>'s text. A covered cell is the slot of a span.
/// </summary>
internal sealed class OdtTableReader(OdtReadContext context, OdtBodyReader body)
{
    /// <summary>Against <c>number-columns-repeated="1000000"</c>.</summary>
    private const int MaxRepeat = 1024;

    private const string NoBorder = "none";

    private const double MinBorderPt = 0.25;

    private static readonly string[] Sides = ["top", "right", "bottom", "left"];

    public IEnumerable<Node> Read(XElement table, bool flat)
    {
        var style = context.Styles.Find("table", table.Attr(Ns.Table + "style-name"));
        var properties = style?.Properties("table-properties");
        var breakBefore = !flat && !context.IsFirstBlock && properties?.Attr(Ns.Fo + "break-before") == ParagraphFormat.Page;
        context.SeeBlock(style?.MasterPage);

        var widths = ColumnWidths(table);
        var rows = new List<Node>();
        foreach (var (row, header) in RowsOf(table, header: false)) rows.Add(Row(row, header, widths));
        if (rows.Count == 0) yield break;

        if (breakBefore) yield return Node.Of("pageBreak");
        var node = Node.Of("table");
        node.Content = rows;
        yield return node;
    }

    /// <summary>In pixels per column, or <c>null</c> for one the file does not measure.</summary>
    private List<int?> ColumnWidths(XElement table)
    {
        var widths = new List<int?>();
        foreach (var column in Columns(table))
        {
            var style = context.Styles.Find("table-column", column.Attr(Ns.Table + "style-name"));
            var width = OdtLength.Pixels(style?.Properties("table-column-properties")?.Attr(Ns.Style + "column-width"));
            var repeat = Math.Clamp(OdtLength.Integer(column.Attr(Ns.Table + "number-columns-repeated")) ?? 1, 1, MaxRepeat);
            for (var index = 0; index < repeat && widths.Count < MaxRepeat; index++) widths.Add(width);
        }

        return widths;
    }

    private static IEnumerable<XElement> Columns(XElement container)
    {
        foreach (var child in container.Elements())
        {
            if (child.Name == Ns.Table + "table-column") yield return child;
            else if (child.Name.LocalName is "table-columns" or "table-header-columns" or "table-column-group")
            {
                foreach (var column in Columns(child)) yield return column;
            }
        }
    }

    /// <summary>Header rows are the editor's <c>tableHeader</c> cells.</summary>
    private static IEnumerable<(XElement Row, bool Header)> RowsOf(XElement container, bool header)
    {
        foreach (var child in container.Elements())
        {
            if (child.Name == Ns.Table + "table-row") yield return (child, header);
            else if (child.Name.LocalName is "table-rows" or "table-row-group" or "table-header-rows")
            {
                foreach (var row in RowsOf(child, header || child.Name.LocalName == "table-header-rows")) yield return row;
            }
        }
    }

    private Node Row(XElement row, bool header, List<int?> widths)
    {
        var cells = new List<Node>();
        var column = 0;
        foreach (var cell in row.Elements())
        {
            var repeat = Math.Clamp(OdtLength.Integer(cell.Attr(Ns.Table + "number-columns-repeated")) ?? 1, 1, MaxRepeat);
            if (cell.Name == Ns.Table + "covered-table-cell")
            {
                column += repeat;
                continue;
            }

            if (cell.Name != Ns.Table + "table-cell") continue;

            // Every column has its element: the columns a span covers come as covered cells.
            for (var index = 0; index < repeat; index++) cells.Add(Cell(cell, header, widths, column++));
        }

        var rowNode = Node.Of("tableRow");
        rowNode.Content = cells.Count == 0 ? [Node.Of(header ? "tableHeader" : "tableCell", OdtParagraphReader.Empty())] : cells;
        return rowNode;
    }

    private Node Cell(XElement cell, bool header, List<int?> widths, int column)
    {
        var content = new List<Node>();
        foreach (var child in cell.Elements()) content.AddRange(body.ReadBlock(child, flat: true));
        if (content.Count == 0) content.Add(OdtParagraphReader.Empty());

        var node = Node.Of(header ? "tableHeader" : "tableCell");
        node.Content = content;
        var colspan = Math.Clamp(OdtLength.Integer(cell.Attr(Ns.Table + "number-columns-spanned")) ?? 1, 1, MaxRepeat);
        node.With("colspan", colspan);
        node.With("rowspan", Math.Clamp(OdtLength.Integer(cell.Attr(Ns.Table + "number-rows-spanned")) ?? 1, 1, MaxRepeat));

        var spanned = widths.Skip(column).Take(colspan).ToList();
        if (spanned.Count == colspan && spanned.All(width => width is > 0))
        {
            node.With("colwidth", new JsonArray([.. spanned.Select(width => (JsonNode?)JsonValue.Create(width!.Value))]));
        }

        var properties = context.Styles.Find("table-cell", cell.Attr(Ns.Table + "style-name"))?.Properties("table-cell-properties");
        if (CharacterFormat.ColorOf(properties?.Attr(Ns.Fo + "background-color")) is { } shading) node.With("shading", shading);
        if (Borders(properties) is { } borders) node.With("borders", borders);
        return node;
    }

    /// <summary>
    /// <c>top:double,1.5,#ff0000;…</c>. A side with the default thin black line says nothing: that
    /// is what the editor draws anyway. A side the cell does not state has no line in ODF, so it
    /// becomes <c>none</c>, or the writer would give it the default one.
    /// </summary>
    private static string? Borders(XElement? properties)
    {
        var text = new StringBuilder();
        foreach (var side in Sides)
        {
            var value = properties?.Attr(Ns.Fo + "border-" + side) ?? properties?.Attr(Ns.Fo + "border") ?? NoBorder;
            if (value.Trim() == OdtStyles.DefaultBorder || Describe(value) is not { } border) continue;
            if (text.Length > 0) text.Append(';');
            text.Append(side).Append(':').Append(border);
        }

        return text.Length == 0 ? null : text.ToString();
    }

    /// <summary><c>0.5pt solid #000000</c> → <c>single,0.5,#000000</c>.</summary>
    private static string? Describe(string value)
    {
        var parts = value.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) return null;
        if (parts.Length == 1 && parts[0] is "none" or "hidden") return $"none,{Format(MinBorderPt)},#000000";

        var width = parts.Select(OdtLength.ExactPoints).FirstOrDefault(points => points is not null) ?? MinBorderPt;
        var color = parts.Select(CharacterFormat.ColorOf).FirstOrDefault(found => found is not null) ?? "#000000";
        var style = parts.FirstOrDefault(part => part is "solid" or "double" or "dotted" or "dashed" or "none" or "hidden");
        var kind = style switch
        {
            "double" => "double",
            "dotted" => "dotted",
            "dashed" => "dashed",
            "none" or "hidden" => "none",
            _ => "single",
        };
        return $"{kind},{Format(Math.Max(MinBorderPt, width))},{color}";
    }

    private static string Format(double points) => points.ToString("0.##", CultureInfo.InvariantCulture);
}

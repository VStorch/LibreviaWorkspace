using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Table → <c>w:tbl</c>, **by position**: each file <c>w:tblPr</c>, <c>w:tblGrid</c>, <c>w:trPr</c>
/// and <c>w:tcPr</c> goes back in place, and only paragraphs are rewritten. A row or column at the
/// end gets new structure; inserting in the middle or removing breaks the correspondence, and that
/// goes to the inventory.
/// </summary>
internal sealed class TableWriter(
    Inventory inventory,
    Func<Node, OpenXmlElement?, IEnumerable<OpenXmlElement>> writeBlock,
    int usableWidthPx,
    bool revisions = true,
    bool verticalMerges = true)
{
    /// <summary>
    /// Removed, a row takes width, merging and shading along; inserted in the middle, it shifts the
    /// structure, and a misplaced <c>w:vMerge</c> makes Word report a corrupt table.
    /// </summary>
    private const string ShiftedStructure =
        "sidecar.notice.shiftedTablePart";

    private readonly TableGridWriter _grid = new(inventory, usableWidthPx);

    public Table Write(Node node, Table? original)
    {
        var table = new Table();

        // Before the first row: properties, grid and the table revision.
        var interleaved = Interleaved<TableRow>(original);
        foreach (var setting in Strays(interleaved, -1)) table.AppendChild(setting.CloneNode(true));
        if (original is null) table.AppendChild(DefaultProperties());

        var originalRows = original?.Elements<TableRow>().ToList() ?? [];
        var rows = node.Content ?? [];

        // The model carries no row identity: with different counts, there is no knowing where it
        // changed.
        var aligned = original is null || originalRows.Count == rows.Count;
        if (!aligned) inventory.NoteLoss(ShiftedStructure);

        var coverage = new TableCoverage();
        // An old draft read every `w:vMerge` as an ordinary cell; see BodyReader.
        var merges = verticalMerges ? VerticalMerges.Of(original) : VerticalMerges.None;
        for (var index = 0; index < rows.Count; index++)
        {
            var row = new RowSource(originalRows.ElementAtOrDefault(index), merges, aligned);
            var slots = coverage.Place(rows[index].Content ?? [], VerticalMerges.GridBeforeOf(row.Row));
            table.AppendChild(WriteRow(rows[index], slots, row));
            foreach (var between in Strays(interleaved, index)) table.AppendChild(between.CloneNode(true));
        }

        // What came after a row that no longer exists closes the table.
        for (var index = rows.Count; index < originalRows.Count; index++)
        {
            foreach (var orphan in Strays(interleaved, index)) table.AppendChild(orphan.CloneNode(true));
        }

        _grid.Apply(table, rows, original);
        return table;
    }

    /// <summary>
    /// Children that are not <typeparamref name="T"/>, by position: <c>-1</c> before the first,
    /// <c>n</c> after the one at index <c>n</c>. Bookmarks, <c>w:sdt</c> and revisions live between
    /// rows, and cannot all go to the start.
    /// </summary>
    private static Dictionary<int, List<OpenXmlElement>> Interleaved<T>(OpenXmlElement? parent)
        where T : OpenXmlElement
    {
        var map = new Dictionary<int, List<OpenXmlElement>>();
        if (parent is null) return map;

        var position = -1;
        foreach (var child in parent.ChildElements)
        {
            if (child is T)
            {
                position++;
                continue;
            }

            if (!map.TryGetValue(position, out var bucket)) map[position] = bucket = [];
            bucket.Add(child);
        }

        return map;
    }

    private static List<OpenXmlElement> Strays(Dictionary<int, List<OpenXmlElement>> map, int position) =>
        map.TryGetValue(position, out var found) ? found : [];

    /// <summary>
    /// A table created here; width and <c>w:tblLayout</c> come with the grid (<see
    /// cref="TableGridWriter"/>).
    /// </summary>
    private static TableProperties DefaultProperties() => new(
        // In schema order, or Word refuses the document.
        new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));

    /// <param name="Aligned">False: this row's structure may belong to another.</param>
    private sealed record RowSource(TableRow? Row, VerticalMerges Merges, bool Aligned);

    private TableRow WriteRow(Node rowNode, List<TableSlot> slots, RowSource source)
    {
        var row = new TableRow();
        var original = source.Row;

        // `w:trPr` carries the height and `w:tblHeader`; `w:tblPrEx`, the row's exceptions.
        var interleaved = Interleaved<TableCell>(original);
        foreach (var setting in Strays(interleaved, -1)) row.AppendChild(setting.CloneNode(true));

        var originalCells = original?.Elements<TableCell>().ToList() ?? [];
        var cells = rowNode.Content ?? [];

        var aligned = original is null || originalCells.Count == slots.Count;
        if (!aligned) inventory.NoteLoss(ShiftedStructure);

        for (var index = 0; index < slots.Count; index++)
        {
            var slot = slots[index];
            var cell = new CellSource(originalCells.ElementAtOrDefault(index), aligned && source.Aligned, source.Merges);
            row.AppendChild(slot.IsContinuation ? Continuation(slot.Span, cell) : WriteCell(slot.Cell!, cell));
            foreach (var between in Strays(interleaved, index)) row.AppendChild(between.CloneNode(true));
        }

        for (var index = slots.Count; index < originalCells.Count; index++)
        {
            foreach (var orphan in Strays(interleaved, index)) row.AppendChild(orphan.CloneNode(true));
        }

        // A row made entirely of `tableHeader` is `w:tblHeader`.
        ApplyHeader(row, cells.Count > 0 && cells.All(cell => cell.Type == "tableHeader"));
        if (revisions) ApplyRowRevision(row, Attr.Node(rowNode, "rowRevision"));
        return row;
    }

    /// <c>w:ins</c>/<c>w:del</c> at the end of <c>w:trPr</c>, only before <c>w:trPrChange</c>.
    private static void ApplyRowRevision(TableRow row, System.Text.Json.Nodes.JsonNode? wanted)
    {
        var properties = row.TableRowProperties;
        if (properties is null)
        {
            if (wanted is null) return;
            properties = new TableRowProperties();
            row.PrependChild(properties);
        }

        Revisions.ApplyBlock(properties, wanted, created =>
        {
            var change = properties.GetFirstChild<TableRowPropertiesChange>();
            if (change is null) properties.AppendChild(created);
            else properties.InsertBefore(created, change);
        });

        if (!properties.HasChildren) properties.Remove();
    }

    /// <summary>
    /// Only when the state changes: <c>w:trPr</c> also carries height and revision.
    /// </summary>
    private static void ApplyHeader(TableRow row, bool wanted)
    {
        var properties = row.TableRowProperties;
        var current = properties?.GetFirstChild<TableHeader>();
        // Present without `w:val` already means "yes".
        var declared = current is not null && !IsOff(current);
        if (declared == wanted) return;

        current?.Remove();
        if (!wanted) return;

        if (properties is null)
        {
            properties = new TableRowProperties();
            row.TableRowProperties = properties;
        }

        // Before the revision, which closes `w:trPr` in the schema.
        var revision = properties.ChildElements.FirstOrDefault(child =>
            child is Inserted or Deleted or TableRowPropertiesChange);
        if (revision is null) properties.AppendChild(new TableHeader());
        else properties.InsertBefore(new TableHeader(), revision);
    }

    private static bool IsOff(TableHeader header) =>
        header.Val is { } value && value.Value == OnOffOnlyValues.Off;

    private static List<OpenXmlElement> ChildrenOf(OpenXmlElement? element) =>
        element is null ? [] : [.. element.ChildElements];

    /// <param name="Aligned">Whether this really is the file's cell at this position.</param>
    private sealed record CellSource(TableCell? Cell, bool Aligned, VerticalMerges Merges)
    {
        public bool WasMerged => Cell is not null && Merges.IsMerged(Cell);

        public bool WasContinuation => Aligned && Cell is not null && Merges.IsAbsorbed(Cell);
    }

    /// <summary>
    /// The lower part of a vertical merge. The file's own continuation goes back as it was: Word keeps
    /// its text, hidden.
    /// </summary>
    private static TableCell Continuation(int span, CellSource source)
    {
        if (source.WasContinuation)
        {
            var kept = (TableCell)source.Cell!.CloneNode(true);
            kept.TableCellProperties!.GridSpan = span > 1 ? new GridSpan { Val = span } : null;
            return kept;
        }

        var properties = new TableCellProperties { VerticalMerge = new VerticalMerge() };
        if (span > 1) properties.GridSpan = new GridSpan { Val = span };
        return new TableCell(properties, new Paragraph());
    }

    private TableCell WriteCell(Node cellNode, CellSource source)
    {
        var cell = new TableCell();
        var original = source.Cell;
        var aligned = source.Aligned;

        var properties = original?.TableCellProperties?.CloneNode(true) as TableCellProperties
                         ?? new TableCellProperties();
        var span = Attr.Int(cellNode, "colspan");

        // Only what the model represents (`colspan`, `rowspan`, shading, borders): cell margins and
        // vertical alignment stay.
        if (span is > 1) properties.GridSpan = new GridSpan { Val = span };
        else if (properties.GridSpan is not null) properties.GridSpan = null;

        // A shifted `w:vMerge` makes Word report a corrupt table. A `w:vMerge` the model does not
        // read as a merge (alone, or with nothing above) stays where it was.
        if (TableCoverage.RowSpanOf(cellNode) > 1)
        {
            properties.VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
        }
        else if (!aligned || source.WasMerged)
        {
            properties.VerticalMerge = null;
        }

        // Only when they differ from the file; see TableLook.
        TableLook.ApplyShading(properties, Attr.String(cellNode, "shading"), inventory);
        TableLook.ApplyBorders(properties, Attr.String(cellNode, "borders"), inventory);

        if (properties.HasChildren) cell.TableCellProperties = properties;

        var children = ChildrenOf(original);
        var blocks = children.Where(child => child is Paragraph or Table).ToList();

        // A content control or a field between paragraphs would come back as nothing.
        if (children.Any(child => child is not (Paragraph or Table or TableCellProperties)))
        {
            inventory.NoteLoss("sidecar.notice.editedCellSpecialContent");
        }

        var position = 0;
        foreach (var child in cellNode.Content ?? [])
        {
            var block = blocks.ElementAtOrDefault(position);
            position++;

            foreach (var element in writeBlock(child, block)) cell.AppendChild(element);
        }

        // A `w:tc` must **end** in a `w:p`, and a nested table may end in a `w:tbl`.
        if (cell.LastChild is not Paragraph) cell.AppendChild(new Paragraph());
        return cell;
    }
}

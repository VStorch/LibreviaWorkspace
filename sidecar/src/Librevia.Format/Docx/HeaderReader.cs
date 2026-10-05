using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Header and footer → a three-column band. The content is usually a **group of shapes**; instead
/// of positioning each piece, it goes to the third it falls in across the width, which is what a
/// corporate header is: left, middle and logo on the right.
/// </summary>
public static class HeaderReader
{
    /// <summary>A shape without text, wide and low: a rule, not a box.</summary>
    private const double RuleAspectRatio = 20;

    /// <summary>
    /// By <c>w:type</c>: in XML order, a document with a title page could show the title page
    /// header on every page.
    /// </summary>
    public static BandDto Read(
        SectionProperties section,
        MainDocumentPart part,
        Inventory inventory,
        HeaderFooterValues type,
        double contentWidthEmus) =>
        ReadReferenced(
            section.Elements<HeaderReference>().Where(r => Matches(r.Type, type)).Select(r => r.Id?.Value),
            part,
            inventory,
            contentWidthEmus);

    public static BandDto ReadFooter(
        SectionProperties section,
        MainDocumentPart part,
        Inventory inventory,
        HeaderFooterValues type,
        double contentWidthEmus) =>
        ReadReferenced(
            section.Elements<FooterReference>().Where(r => Matches(r.Type, type)).Select(r => r.Id?.Value),
            part,
            inventory,
            contentWidthEmus);

    /// <summary>Without <c>w:type</c> it is <c>default</c>, as the specification says.</summary>
    private static bool Matches(EnumValue<HeaderFooterValues>? declared, HeaderFooterValues wanted) =>
        (declared?.Value ?? HeaderFooterValues.Default) == wanted;

    private static BandDto ReadReferenced(
        IEnumerable<string?> relationshipIds,
        MainDocumentPart part,
        Inventory inventory,
        double contentWidthEmus)
    {
        foreach (var id in relationshipIds)
        {
            if (string.IsNullOrEmpty(id)) continue;

            OpenXmlPartRootElement? root;
            OpenXmlPart? owner;

            switch (part.GetPartById(id))
            {
                case HeaderPart header:
                    root = header.Header;
                    owner = header;
                    break;
                case FooterPart footer:
                    root = footer.Footer;
                    owner = footer;
                    break;
                default:
                    continue;
            }

            if (root is null || owner is null) continue;

            var band = Build(root, owner, inventory, contentWidthEmus, new FontTable(part), id);
            // More than one reference of the same type: the one with content wins.
            if (!band.IsEmpty) return band;
        }

        return BandDto.Empty();
    }

    private static BandDto Build(
        OpenXmlPartRootElement root,
        OpenXmlPart owner,
        Inventory inventory,
        double contentWidthEmus,
        FontTable fonts,
        string relationshipId)
    {
        // Each paragraph's address: editing finds the `w:t` to write by it.
        var naming = new Naming(relationshipId, BandNav.IndexOf(root));

        // The same for boxes, where the corporate header title lives.
        var boxes = new Boxing(relationshipId, BandNav.BoxIndexOf(root));

        var columns = new List<PieceDto>[3];
        for (var i = 0; i < 3; i++) columns[i] = [];

        var floats = new List<FloatDto>();
        var rule = false;

        // The grid first: the thirds passes must not see it twice.
        var rows = ReadGrid(root, owner, inventory, fonts, naming);

        foreach (var paragraph in root.Descendants<Paragraph>())
        {
            // Box paragraphs go with the shape, to inherit its position.
            if (paragraph.Ancestors<TextBoxContent>().Any()) continue;
            if (paragraph.Ancestors<Table>().Any()) continue;

            if (HasBottomBorder(paragraph)) rule = true;

            var pieces = ReadRuns(paragraph, inventory, fonts, naming);
            if (pieces.Count == 0) continue;

            var column = columns[ColumnOf(paragraph)];
            column.AddRange(OpeningALine(pieces, column.Count > 0));
        }

        foreach (var drawing in root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
        {
            // The VML fallback repeats the content.
            if (drawing.Ancestors<AlternateContentFallback>().Any()) continue;
            if (drawing.Ancestors<Table>().Any()) continue;

            // An anchored object has a real position and may come rotated: it goes out as an
            // object, not as a column piece.
            if (AnchorReader.AnchorOf(drawing) is { } anchor)
            {
                foreach (var item in ReadAnchoredDrawing(drawing, owner, anchor, inventory, fonts, boxes))
                {
                    floats.Add(item);
                }
                continue;
            }

            ReadDrawing(drawing, owner, columns, ref rule, inventory, contentWidthEmus, fonts, naming);
        }

        return new BandDto(columns[0], columns[1], columns[2], rule, floats, rows);
    }


    /// <summary>
    /// OOXML vertical merging is <c>restart</c> on top and an empty <c>w:vMerge</c> below: here the
    /// lower ones disappear and the top one grows, as HTML understands it.
    /// </summary>
    private static List<BandRowDto> ReadGrid(
        OpenXmlPartRootElement root,
        OpenXmlPart owner,
        Inventory inventory,
        FontTable fonts,
        Naming naming)
    {
        var rows = new List<BandRowDto>();

        foreach (var table in root.Elements<Table>())
        {
            var borders = table.GetFirstChild<TableProperties>()?.TableBorders;
            var total = ColumnWidths(table).Sum();
            if (total <= 0) total = 1;

            // The open cell in each column, so a continuation knows whom to add to.
            var open = new Dictionary<int, BandCellDto>();
            var built = new List<List<BandCellDto>>();
            var trs = table.Elements<TableRow>().ToList();

            for (var r = 0; r < trs.Count; r++)
            {
                var cells = new List<BandCellDto>();
                var column = 0;

                foreach (var cell in trs[r].Elements<TableCell>())
                {
                    var properties = cell.TableCellProperties;
                    var span = properties?.GridSpan?.Val?.Value ?? 1;
                    var merge = properties?.VerticalMerge;

                    // `w:vMerge` without `w:val` or with `continue` is the lower cell.
                    if (merge is not null && merge.Val?.Value != MergedCellValues.Restart)
                    {
                        if (open.TryGetValue(column, out var above))
                        {
                            open[column] = above with { RowSpan = above.RowSpan + 1 };
                            ReplaceIn(built, above, open[column]);
                        }

                        column += span;
                        continue;
                    }

                    var width = TwipsOf(properties?.TableCellWidth) / total;
                    var built_ = new BandCellDto(
                        ReadCellPieces(cell, owner, inventory, fonts, naming),
                        Math.Round(width, 4),
                        span,
                        1,
                        AlignOf(cell),
                        BordersOf(cell, borders, r == 0, r == trs.Count - 1, column == 0));

                    cells.Add(built_);
                    if (merge is not null) open[column] = built_;
                    column += span;
                }

                built.Add(cells);
            }

            foreach (var row in built)
            {
                if (row.Count > 0) rows.Add(new BandRowDto(row));
            }
        }

        return rows;
    }

    private static void ReplaceIn(List<List<BandCellDto>> built, BandCellDto old, BandCellDto grown)
    {
        foreach (var row in built)
        {
            var at = row.IndexOf(old);
            if (at >= 0)
            {
                row[at] = grown;
                return;
            }
        }
    }

    private static IEnumerable<double> ColumnWidths(Table table) =>
        table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
            .Select(column => double.TryParse(column.Width?.Value, out var width) ? width : 0)
        ?? [];

    private static double TwipsOf(TableCellWidth? width) =>
        double.TryParse(width?.Width?.Value, out var value) ? value : 0;

    private static string? AlignOf(TableCell cell)
    {
        var value = cell.Descendants<Paragraph>()
            .Select(paragraph => paragraph.ParagraphProperties?.Justification?.Val)
            .FirstOrDefault(justification => justification is not null);

        if (value is null) return null;
        if (value == JustificationValues.Center) return "center";
        if (value == JustificationValues.Right) return "right";
        return null;
    }

    /// <summary>
    /// The cell's border, the table's outer one on the outer side, the inner one on the other;
    /// <c>nil</c> on the cell erases the table's line.
    /// </summary>
    private static string BordersOf(
        TableCell cell,
        TableBorders? table,
        bool firstRow,
        bool lastRow,
        bool firstColumn)
    {
        var own = cell.TableCellProperties?.TableCellBorders;
        var sides = string.Empty;

        if (Drawn(own?.TopBorder, firstRow ? table?.TopBorder : table?.InsideHorizontalBorder)) sides += "t";
        if (Drawn(own?.LeftBorder, firstColumn ? table?.LeftBorder : table?.InsideVerticalBorder)) sides += "l";
        if (Drawn(own?.BottomBorder, lastRow ? table?.BottomBorder : table?.InsideHorizontalBorder)) sides += "b";
        // The right side uses the outer one: when in doubt, draw too much rather than too little.
        if (Drawn(own?.RightBorder, table?.RightBorder)) sides += "r";

        return sides;
    }

    private static bool Drawn(BorderType? own, BorderType? inherited)
    {
        var border = own ?? inherited;
        if (border?.Val?.Value is not { } style) return false;
        return style != BorderValues.None && style != BorderValues.Nil;
    }

    /// <summary>Images and text, in that order.</summary>
    private static List<PieceDto> ReadCellPieces(
        TableCell cell,
        OpenXmlPart owner,
        Inventory inventory,
        FontTable fonts,
        Naming naming)
    {
        var pieces = new List<PieceDto>();

        foreach (var paragraph in cell.Descendants<Paragraph>())
        {
            foreach (var picture in paragraph.Descendants<Drawing.Pictures.Picture>())
            {
                if (ImagePieceOf(picture, owner) is { } image) pieces.Add(image);
            }

            var run = ReadRuns(paragraph, inventory, fonts, naming);
            if (run.Count > 0) pieces.AddRange(OpeningALine(run, pieces.Count > 0));
        }

        return pieces;
    }

    /// <summary>The first piece opens a line only when something already precedes it.</summary>
    private static IEnumerable<PieceDto> OpeningALine(List<PieceDto> pieces, bool after)
    {
        for (var index = 0; index < pieces.Count; index++)
        {
            yield return index == 0 && after ? pieces[index] with { Line = true } : pieces[index];
        }
    }

    private static PieceDto? ImagePieceOf(Drawing.Pictures.Picture picture, OpenXmlPart owner)
    {
        var relationshipId = picture.Descendants<Drawing.Blip>().FirstOrDefault()?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;
        if (owner.GetPartById(relationshipId) is not ImagePart image) return null;

        var extent = picture.Ancestors<OpenXmlElement>()
            .SelectMany(element => element.Elements<WordDrawing.Extent>())
            .FirstOrDefault();
        var (_, width, height) = GeometryOf(picture.Descendants<Drawing.Transform2D>().FirstOrDefault());
        if (width <= 0) width = (double?)extent?.Cx?.Value ?? 0;
        if (height <= 0) height = (double?)extent?.Cy?.Value ?? 0;
        if (width <= 0) return null;

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);

        return PieceDto.Image(
            $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}",
            Pixels(width),
            Pixels(height > 0 ? height : width / 4));
    }

    /// <summary>
    /// Each group piece in its own box: the group box would stretch the logo, and looking only for
    /// images would drop the header title.
    /// </summary>
    private static IEnumerable<FloatDto> ReadAnchoredDrawing(
        OpenXmlElement drawing,
        OpenXmlPart owner,
        WordDrawing.Anchor anchor,
        Inventory inventory,
        FontTable fonts,
        Boxing boxes)
    {
        foreach (var piece in AnchorReader.PiecesOf(anchor))
        {
            if (piece.Shape is Drawing.Pictures.Picture picture)
            {
                if (ImageSourceOf(picture, owner) is { } src)
                {
                    yield return AnchorReader.Describe(anchor, "image", src, null, piece);
                }

                continue;
            }

            var content = new List<Node>();
            var inside = BandNav.BoxesOf(piece.Shape);

            foreach (var box in inside)
            {
                foreach (var paragraph in box.Descendants<Paragraph>())
                {
                    var pieces = ReadRuns(paragraph, inventory, fonts);
                    if (pieces.Count == 0) continue;

                    var node = Node.Of("paragraph");
                    node.Content = pieces.Select(NodeOf).ToList();
                    content.Add(node);
                }
            }

            if (content.Count > 0)
            {
                // Only a single-box shape is editable: with two, there would be no knowing where
                // the text goes back.
                var address = inside.Count == 1 ? boxes.Of(inside[0]) : null;

                var look = ShapeLook.Of(piece.Shape);
                if (!look.Complete) inventory.NoteInvisible(Inventory.Shapes);

                yield return AnchorReader.Describe(anchor, "text", null, content, piece) with
                {
                    BoxId = address,
                    Fill = look.Fill,
                    Line = look.Line,
                    LineWidthPt = look.LineWidthPt,
                    Dash = look.Dashed,
                };
                continue;
            }

            // The rule under the header: a pair of zero-height shapes in the logo group.
            if (IsRule(piece)) yield return AnchorReader.Describe(anchor, "rule", null, null, piece);
        }

        if (AnchorReader.PiecesOf(anchor).Count > 0) yield break;

        foreach (var picture in drawing.Descendants<Drawing.Pictures.Picture>())
        {
            if (ImageSourceOf(picture, owner) is { } src)
            {
                yield return AnchorReader.Describe(anchor, "image", src, null);
            }
        }
    }

    /// <summary>
    /// As the screen shows it, with <c>PAGE</c> as <c>{n}</c>: compared with the XML, the box would
    /// always differ and the field would become a literal <c>{n}</c>.
    /// </summary>
    internal static string BoxTextOf(TextBoxContent box, Inventory inventory, FontTable fonts) =>
        string.Join(
            "\n",
            box.Descendants<Paragraph>()
                .Select(paragraph => ReadRuns(paragraph, inventory, fonts))
                .Where(pieces => pieces.Count > 0)
                .Select(pieces => string.Concat(pieces.Select(TextOf))));

    private static string TextOf(PieceDto piece) => piece.Kind switch
    {
        PieceDto.KindPageNumber => "{n}",
        PieceDto.KindTotalPages => "{total}",
        _ => piece.Text ?? string.Empty,
    };

    private static Node NodeOf(PieceDto piece)
    {
        var text = TextOf(piece);

        var marks = new List<Mark>();
        if (piece.Bold) marks.Add(Mark.Of("bold"));
        if (piece.Italic) marks.Add(Mark.Of("italic"));

        var attributes = new Dictionary<string, System.Text.Json.Nodes.JsonNode?>();
        if (piece.Color is not null) attributes["color"] = piece.Color;
        if (piece.FontSize is not null) attributes["fontSize"] = piece.FontSize;
        if (piece.FontFamily is not null) attributes["fontFamily"] = piece.FontFamily;
        if (attributes.Count > 0) marks.Add(new Mark { Type = "textStyle", Attrs = attributes });

        return new Node
        {
            Type = "text",
            Text = text,
            Marks = marks.Count == 0 ? null : marks,
        };
    }

    /// <summary>
    /// Wide, flat, zero-height and outlined: without a stroke it is reserved space.
    /// </summary>
    private static bool IsRule(AnchoredPiece piece)
    {
        if (piece.WidthEmus <= 0) return false;
        if (piece.HeightEmus > 0 && piece.WidthEmus / piece.HeightEmus < RuleAspectRatio) return false;

        var outline = piece.Shape.Descendants<Drawing.Outline>().FirstOrDefault();
        return outline is not null && (outline.Width?.Value ?? 0) > 0;
    }

    private static string? ImageSourceOf(OpenXmlElement picture, OpenXmlPart owner)
    {
        var relationshipId = picture.Descendants<Drawing.Blip>().FirstOrDefault()?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;
        if (owner.GetPartById(relationshipId) is not ImagePart image) return null;

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);

        return $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}";
    }

    private static void ReadDrawing(
        OpenXmlElement drawing,
        OpenXmlPart owner,
        List<PieceDto>[] columns,
        ref bool rule,
        Inventory inventory,
        double contentWidthEmus,
        FontTable fonts,
        Naming naming)
    {
        // The position on the page comes from the anchor: the `a:off` of a single-piece drawing is
        // zero.
        var anchor = drawing.Descendants<WordDrawing.Anchor>().FirstOrDefault();
        var horizontal = anchor?.GetFirstChild<WordDrawing.HorizontalPosition>();
        var anchorOffset = long.TryParse(horizontal?.PositionOffset?.Text, out var emus)
            ? (double?)emus
            : null;
        var anchorAlign = horizontal?.HorizontalAlignment?.Text;

        var totalWidth = (double?)drawing.Descendants<WordDrawing.Extent>().FirstOrDefault()?.Cx?.Value;
        if (totalWidth is null or <= 0) totalWidth = 1;

        var groupExtent = drawing.Descendants<Drawing.ChildExtents>().FirstOrDefault();
        var span = (double?)groupExtent?.Cx?.Value ?? totalWidth.Value;
        var origin = (double?)drawing.Descendants<Drawing.ChildOffset>().FirstOrDefault()?.X?.Value ?? 0;

        foreach (var shape in drawing.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape>())
        {
            var (offset, width, height) = GeometryOf(shape.Descendants<Drawing.Transform2D>().FirstOrDefault());
            var pieces = shape.Descendants<TextBoxContent>()
                .SelectMany(box => box.Descendants<Paragraph>())
                .SelectMany(paragraph => ReadRuns(paragraph, inventory, fonts, naming))
                .ToList();

            if (pieces.Count == 0)
            {
                if (height > 0 && width / height >= RuleAspectRatio) rule = true;
                continue;
            }

            columns[ColumnFor(offset, width, origin, span, anchorOffset, anchorAlign, contentWidthEmus)]
                .AddRange(pieces);
        }

        foreach (var picture in drawing.Descendants<Drawing.Pictures.Picture>())
        {
            var relationshipId = picture.Descendants<Drawing.Blip>().FirstOrDefault()?.Embed?.Value;
            if (string.IsNullOrEmpty(relationshipId)) continue;
            if (owner.GetPartById(relationshipId) is not ImagePart image) continue;

            var (offset, width, height) = GeometryOf(picture.Descendants<Drawing.Transform2D>().FirstOrDefault());
            if (width <= 0) width = totalWidth.Value;

            using var stream = image.GetStream();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);

            columns[ColumnFor(offset, width, origin, span, anchorOffset, anchorAlign, contentWidthEmus)]
                .Add(PieceDto.Image(
                $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}",
                Pixels(width),
                Pixels(height > 0 ? height : width / 4)));
        }
    }

    private static (double Offset, double Width, double Height) GeometryOf(Drawing.Transform2D? transform) =>
        (
            (double?)transform?.Offset?.X?.Value ?? 0,
            (double?)transform?.Extents?.Cx?.Value ?? 0,
            (double?)transform?.Extents?.Cy?.Value ?? 0);

    private static int Pixels(double emu) => (int)Math.Round(Unit.EmusToPixels(emu));

    private static int ColumnFor(
        double offset,
        double width,
        double origin,
        double span,
        double? anchorOffset,
        string? anchorAlign,
        double contentWidthEmus)
    {
        // A declared alignment already gives the third.
        if (anchorAlign is not null)
        {
            return anchorAlign switch
            {
                "center" => 1,
                "right" or "outside" => 2,
                _ => 0,
            };
        }

        if (anchorOffset is not null && contentWidthEmus > 0)
        {
            // In a group, the coordinate is in group space: the sum errs inside it, never on the
            // page third.
            return ThirdOf((anchorOffset.Value + offset - origin + width / 2) / contentWidthEmus);
        }

        if (span <= 0) return 0;
        return ThirdOf((offset - origin + width / 2) / span);
    }

    private static int ThirdOf(double fraction) =>
        fraction < 1.0 / 3 ? 0 : fraction < 2.0 / 3 ? 1 : 2;

    private static int ColumnOf(Paragraph paragraph)
    {
        var value = paragraph.ParagraphProperties?.Justification?.Val;
        if (value is null) return 0;
        if (value == JustificationValues.Center) return 1;
        if (value == JustificationValues.Right) return 2;
        return 0;
    }

    private static bool HasBottomBorder(Paragraph paragraph)
    {
        var border = paragraph.ParagraphProperties?.ParagraphBorders?.BottomBorder;
        return border?.Val is not null && border.Val.Value != BorderValues.None;
    }


    /// <summary>
    /// The cached value between <c>separate</c> and <c>end</c> does not come along: it would give
    /// "{n}5".
    /// </summary>
    private sealed class FieldState
    {
        public bool InCachedResult;
    }

    /// <summary>
    /// The trace is the way back for editing, without redoing the run merge. A piece without
    /// <c>w:t</c> (number, image, tab) is not editable.
    /// </summary>
    internal sealed record TracedPiece(PieceDto Piece, List<Text> Source);

    internal sealed record Naming(string RelationshipId, Dictionary<Paragraph, int> Index)
    {
        public string? Of(Paragraph paragraph, int piece) =>
            Index.TryGetValue(paragraph, out var at) ? BandNav.Address(RelationshipId, at, piece) : null;
    }

    internal sealed record Boxing(string RelationshipId, Dictionary<TextBoxContent, int> Index)
    {
        public string? Of(TextBoxContent box) =>
            Index.TryGetValue(box, out var at) ? BandNav.BoxAddress(RelationshipId, at) : null;
    }

    private static List<PieceDto> ReadRuns(
        Paragraph paragraph,
        Inventory inventory,
        FontTable fonts,
        Naming? naming = null)
    {
        var traced = TracedRuns(paragraph, inventory, fonts);
        var pieces = new List<PieceDto>(traced.Count);

        for (var at = 0; at < traced.Count; at++)
        {
            var piece = traced[at].Piece;

            // Blank pieces are not drawn, but they are dropped here: the address is the position in
            // the walk.
            if (piece.Kind == PieceDto.KindText && string.IsNullOrWhiteSpace(piece.Text)) continue;

            pieces.Add(naming is null || traced[at].Source.Count == 0
                ? piece
                : piece with { Pid = naming.Of(paragraph, at) });
        }

        return pieces;
    }

    internal static List<TracedPiece> TracedRuns(Paragraph paragraph, Inventory inventory, FontTable fonts)
    {
        var pieces = new List<TracedPiece>();
        var field = new FieldState();
        Collect(paragraph, pieces, field, inventory, fonts);

        // Joins neighbouring texts with the same style, which Word chops into runs; an equation in
        // the middle separates them.
        var order = new Dictionary<OpenXmlElement, int>(ReferenceEqualityComparer.Instance);
        foreach (var element in paragraph.Descendants()) order[element] = order.Count;
        var equations = paragraph.Descendants()
            .Where(element => element is DocumentFormat.OpenXml.Math.OfficeMath or DocumentFormat.OpenXml.Math.Paragraph)
            .Select(element => order[element]).ToList();
        bool Apart(List<Text> before, List<Text> after) =>
            equations.Count > 0 && before.Count > 0 && after.Count > 0 &&
            equations.Any(at => at > order[before[^1]] && at < order[after[0]]);

        var merged = new List<TracedPiece>();
        foreach (var traced in pieces)
        {
            var piece = traced.Piece;
            var previous = merged.Count > 0 ? merged[^1].Piece : null;
            if (piece.Kind == PieceDto.KindText && previous is { Kind: PieceDto.KindText } &&
                !Apart(merged[^1].Source, traced.Source) &&
                previous.Bold == piece.Bold && previous.Italic == piece.Italic &&
                previous.Color == piece.Color && previous.FontSize == piece.FontSize &&
                previous.FontFamily == piece.FontFamily)
            {
                // With a piece without a trace (a tab is `w:tab`), merging erases both traces.
                var source = merged[^1].Source.Count == 0 || traced.Source.Count == 0
                    ? new List<Text>()
                    : [.. merged[^1].Source, .. traced.Source];

                merged[^1] = new TracedPiece(previous with { Text = previous.Text + piece.Text }, source);
                continue;
            }

            merged.Add(traced);
        }

        // Text that already has `{n}` written is text, not a field.
        return [.. merged.Select(traced =>
            traced.Piece.Kind == PieceDto.KindText && FieldTokens.Contains(traced.Piece.Text)
                ? traced with { Piece = traced.Piece with { Literal = true } }
                : traced)];
    }

    /// <summary>
    /// By the first word: <c>PAGEREF</c> and <c>SECTIONPAGES</c> also contain "PAGE".
    /// </summary>
    private static string? FieldKindOf(string instruction)
    {
        var word = instruction.Trim().Split((char[]?)null, 2, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return word?.ToUpperInvariant() switch
        {
            "PAGE" => PieceDto.KindPageNumber,
            "NUMPAGES" => PieceDto.KindTotalPages,
            _ => null,
        };
    }

    private static void Collect(
        OpenXmlElement parent,
        List<TracedPiece> pieces,
        FieldState field,
        Inventory inventory,
        FontTable fonts)
    {
        foreach (var element in parent.ChildElements)
        {
            switch (element)
            {
                // Drawings have their own pass, with position.
                case DocumentFormat.OpenXml.Wordprocessing.Drawing:
                case Picture:
                case AlternateContent:
                    break;

                case FieldChar marker:
                    if (marker.FieldCharType?.Value is { } type)
                    {
                        if (type == FieldCharValues.Separate) field.InCachedResult = true;
                        else if (type == FieldCharValues.End) field.InCachedResult = false;
                    }

                    break;

                // Word's "Page Number" `w:fldSimple`: without this case the cached number would
                // come.
                case SimpleField simple:
                {
                    if (FieldKindOf(simple.Instruction?.Value ?? string.Empty) is { } simpleKind)
                    {
                        pieces.Add(new TracedPiece(new PieceDto(simpleKind), []));
                    }
                    else
                    {
                        Collect(element, pieces, field, inventory, fonts);
                    }

                    break;
                }

                case FieldCode code:
                    if (FieldKindOf(code.Text) is { } codeKind)
                    {
                        pieces.Add(new TracedPiece(new PieceDto(codeKind), []));
                    }
                    else
                    {
                        inventory.NoteInvisible(Inventory.HeaderFields);
                    }

                    break;

                case Run run:
                    Collect(run, pieces, field, StyleOf(run.RunProperties, fonts), inventory, fonts);
                    break;

                // An equation in a band stays in the file, and the warning says it exists.
                case DocumentFormat.OpenXml.Math.OfficeMath:
                case DocumentFormat.OpenXml.Math.Paragraph:
                    inventory.NoteInvisible(Inventory.Equations);
                    break;

                default:
                    Collect(element, pieces, field, inventory, fonts);
                    break;
            }
        }
    }

    private static void Collect(
        Run run,
        List<TracedPiece> pieces,
        FieldState field,
        PieceDto style,
        Inventory inventory,
        FontTable fonts)
    {
        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case Text text when !field.InCachedResult:
                    pieces.Add(new TracedPiece(style with { Text = text.Text }, [text]));
                    break;

                case TabChar when !field.InCachedResult:
                    pieces.Add(new TracedPiece(style with { Text = " " }, []));
                    break;

                case FieldChar or FieldCode:
                    Collect(run, pieces, field, inventory, fonts);
                    return;

                case RunProperties:
                case Text:
                case TabChar:
                    break;
            }
        }
    }

    private static PieceDto StyleOf(RunProperties? properties, FontTable fonts) => new(
        PieceDto.KindText,
        Bold: RunReader.IsOn(properties?.Bold),
        Italic: RunReader.IsOn(properties?.Italic),
        Color: ColorOf(properties?.Color?.Val?.Value),
        FontSize: SizeOf(properties?.FontSize?.Val?.Value),
        FontFamily: FamilyOf(properties, fonts));

    /// <summary>Without it, the header would inherit the editor font.</summary>
    private static string? FamilyOf(RunProperties? properties, FontTable fonts)
    {
        var font = properties?.RunFonts?.Ascii?.Value ?? properties?.RunFonts?.HighAnsi?.Value;
        return string.IsNullOrWhiteSpace(font) ? null : fonts.Stack(font);
    }

    private static string? ColorOf(string? value) =>
        string.IsNullOrWhiteSpace(value) || value.Equals("auto", StringComparison.OrdinalIgnoreCase)
            ? null
            : "#" + value.TrimStart('#').ToLowerInvariant();

    private static string? SizeOf(string? halfPoints) =>
        double.TryParse(halfPoints, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? Unit.HalfPointsToPoints(value).ToString("0.#", CultureInfo.InvariantCulture) + "pt"
            : null;
}

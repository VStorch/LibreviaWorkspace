using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Cabeçalho e rodapé → faixa de três colunas. O conteúdo costuma ser um **grupo
/// de formas**; em vez de posicionar cada peça, ela vai para o terço em que cai
/// na largura, que é o que o cabeçalho corporativo é: esquerda, meio e logotipo à
/// direita.
/// </summary>
public static class HeaderReader
{
    /// <summary>Forma sem texto, larga e baixa: é filete, não caixa.</summary>
    private const double RuleAspectRatio = 20;

    /// <summary>
    /// Pelo <c>w:type</c>: na ordem do XML, o documento com capa podia exibir o
    /// cabeçalho da capa em todas as páginas.
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

    /// <summary>Sem <c>w:type</c> é <c>default</c>, como diz a especificação.</summary>
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
            // Mais de uma referência do mesmo tipo: vale a que tem conteúdo.
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
        // O endereço de cada parágrafo: é por ele que a edição acha o `w:t` a escrever.
        var naming = new Naming(relationshipId, BandNav.IndexOf(root));

        // O mesmo para as caixas, onde mora o título do cabeçalho corporativo.
        var boxes = new Boxing(relationshipId, BandNav.BoxIndexOf(root));

        var columns = new List<PieceDto>[3];
        for (var i = 0; i < 3; i++) columns[i] = [];

        var floats = new List<FloatDto>();
        var rule = false;

        // A grade primeiro: os passes dos terços não podem vê-la duas vezes.
        var rows = ReadGrid(root, owner, inventory, fonts, naming);

        foreach (var paragraph in root.Descendants<Paragraph>())
        {
            // Os parágrafos da caixa vão com a forma, para herdar a posição.
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
            // O fallback VML repete o conteúdo.
            if (drawing.Ancestors<AlternateContentFallback>().Any()) continue;
            if (drawing.Ancestors<Table>().Any()) continue;

            // O ancorado tem posição de verdade e pode vir girado: sai como objeto, e não como peça de coluna.
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
    /// A mesclagem vertical do OOXML é <c>restart</c> em cima e <c>w:vMerge</c> vazio
    /// embaixo: aqui as de baixo somem e a de cima cresce, como o HTML entende.
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

            // A célula aberta em cada coluna, para a continuação saber a quem somar.
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

                    // `w:vMerge` sem `w:val` ou com `continue` é a célula de baixo.
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
    /// A borda da célula, a externa da tabela no lado externo, a interna no outro;
    /// <c>nil</c> na célula apaga o risco da tabela.
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
        // A direita usa a externa: na dúvida, desenha-se demais, e não de menos.
        if (Drawn(own?.RightBorder, table?.RightBorder)) sides += "r";

        return sides;
    }

    private static bool Drawn(BorderType? own, BorderType? inherited)
    {
        var border = own ?? inherited;
        if (border?.Val?.Value is not { } style) return false;
        return style != BorderValues.None && style != BorderValues.Nil;
    }

    /// <summary>Imagens e texto, nessa ordem.</summary>
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

    /// <summary>A primeira peça abre linha só quando já há algo antes.</summary>
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
    /// Cada peça do grupo na caixa dela: a caixa do grupo esticaria o logotipo, e
    /// procurar só imagens tiraria o título do cabeçalho.
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
                // Só a forma de uma caixa só é editável: com duas, não se saberia para qual voltar o texto.
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

            // O filete sob o cabeçalho: um par de formas de altura zero no grupo do logotipo.
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
    /// Como a tela o mostra, com o <c>PAGE</c> como <c>{n}</c>: comparado com o XML, a
    /// caixa mudaria sempre e o campo viraria <c>{n}</c> literal.
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

    /// <summary>Larga, rasa, de altura zero e com contorno: sem traço é espaço reservado.</summary>
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
        // A posição na página vem da âncora: o `a:off` de um desenho de peça única é zero.
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

    /// <summary>EMU → pixels CSS: 914400 por polegada, 96 px por polegada.</summary>
    private static int Pixels(double emu) => (int)Math.Round(emu * 96 / 914400);

    private static int ColumnFor(
        double offset,
        double width,
        double origin,
        double span,
        double? anchorOffset,
        string? anchorAlign,
        double contentWidthEmus)
    {
        // Alinhamento declarado já diz o terço.
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
            // No grupo, a coordenada é do espaço do grupo: a soma erra dentro dele, nunca o terço da página.
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


    /// <summary>O valor em cache entre <c>separate</c> e <c>end</c> não vai junto: daria "{n}5".</summary>
    private sealed class FieldState
    {
        public bool InCachedResult;
    }

    /// <summary>
    /// O rastro é o caminho de volta da edição, sem refazer a fusão de runs. Peça
    /// sem <c>w:t</c> (número, imagem, tabulação) não é editável.
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

            // Em branco não se desenha, mas o descarte é aqui: o endereço é a posição na travessia.
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

        // Junta textos vizinhos de mesmo estilo, que o Word pica em runs; uma
        // equação no meio separa.
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
                // Com peça sem rastro (a tabulação é `w:tab`), a fusão apaga o rastro das duas.
                var source = merged[^1].Source.Count == 0 || traced.Source.Count == 0
                    ? new List<Text>()
                    : [.. merged[^1].Source, .. traced.Source];

                merged[^1] = new TracedPiece(previous with { Text = previous.Text + piece.Text }, source);
                continue;
            }

            merged.Add(traced);
        }

        // Texto que já traz `{n}` escrito é texto, e não campo.
        return [.. merged.Select(traced =>
            traced.Piece.Kind == PieceDto.KindText && FieldTokens.Contains(traced.Piece.Text)
                ? traced with { Piece = traced.Piece with { Literal = true } }
                : traced)];
    }

    /// <summary>Pela primeira palavra: <c>PAGEREF</c> e <c>SECTIONPAGES</c> também contêm "PAGE".</summary>
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
                // Os desenhos têm passe próprio, com posição.
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

                // O `w:fldSimple` do "Número da página" do Word: sem este caso viria o número em cache.
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

                // A equação da faixa fica no arquivo, e o aviso diz que existe.
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

    /// <summary>Sem ela, o cabeçalho herdaria a fonte do editor.</summary>
    private static string? FamilyOf(RunProperties? properties, FontTable fonts)
    {
        var font = properties?.RunFonts?.Ascii?.Value ?? properties?.RunFonts?.HighAnsi?.Value;
        return string.IsNullOrWhiteSpace(font) ? null : fonts.Stack(font);
    }

    private static string? ColorOf(string? value) =>
        string.IsNullOrWhiteSpace(value) || value.Equals("auto", StringComparison.OrdinalIgnoreCase)
            ? null
            : "#" + value.TrimStart('#').ToLowerInvariant();

    /// <summary><c>w:sz</c> vem em meios-pontos.</summary>
    private static string? SizeOf(string? halfPoints) =>
        double.TryParse(halfPoints, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? (value / 2).ToString("0.#", CultureInfo.InvariantCulture) + "pt"
            : null;
}

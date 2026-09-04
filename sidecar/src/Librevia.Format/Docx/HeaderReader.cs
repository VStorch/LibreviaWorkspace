using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Cabeçalho e rodapé → faixa de três colunas, para desenhar na tela e no PDF.
/// </summary>
/// <remarks>
/// O conteúdo raramente está em parágrafos simples. No corpus real ele vive
/// dentro de um **grupo de formas**: uma imagem de logotipo, uma caixa de texto
/// com o campo `PAGE`, outra com o título e duas formas de decoração. Cada peça
/// tem posição própria dentro do grupo.
///
/// Reproduzir posicionamento absoluto seria caro e frágil. Em vez disso,
/// olhamos **onde cada peça cai na largura** e a jogamos no terço
/// correspondente. O resultado bate com o original nos casos que importam,
/// porque cabeçalho corporativo é quase sempre exatamente isto: algo à
/// esquerda, algo no meio, o logotipo à direita.
/// </remarks>
public static class HeaderReader
{
    /// <summary>Forma sem texto, larga e baixa: é filete, não caixa.</summary>
    private const double RuleAspectRatio = 20;

    /// <summary>
    /// O cabeçalho de um dos três papéis que a seção declara.
    /// </summary>
    /// <remarks>
    /// O `w:type` era ignorado: percorria-se as referências na ordem em que
    /// estavam no XML e ficava a primeira não vazia. Qual cabeçalho aparecia
    /// dependia, portanto, da ordem de gravação — um documento com capa podia
    /// exibir o cabeçalho da capa em todas as páginas. No corpus real, quatro de
    /// seis documentos declaram `first`, `even` e `default`, então o acaso
    /// decidia na maioria dos casos.
    /// </remarks>
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

    /// <summary>
    /// Referência sem `w:type` é `default` — é o que a especificação diz, e é o
    /// que documento antigo grava.
    /// </summary>
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

            var band = Build(root, owner, inventory, contentWidthEmus);
            // Dentro de um mesmo tipo raramente há mais de uma referência; se
            // houver, vale a que tem conteúdo.
            if (!band.IsEmpty) return band;
        }

        return BandDto.Empty();
    }

    private static BandDto Build(
        OpenXmlPartRootElement root,
        OpenXmlPart owner,
        Inventory inventory,
        double contentWidthEmus)
    {
        var columns = new List<PieceDto>[3];
        for (var i = 0; i < 3; i++) columns[i] = [];

        var floats = new List<FloatDto>();
        var rule = false;

        // A grade primeiro: o que está dentro dela tem posição própria, e os dois
        // passes abaixo — os que espalham peças pelos três terços — não podem
        // vê-la duas vezes.
        var rows = ReadGrid(root, owner, inventory);

        foreach (var paragraph in root.Descendants<Paragraph>())
        {
            // Parágrafos de dentro de caixa de texto são tratados junto com a
            // forma que os contém, para herdar a posição dela.
            if (paragraph.Ancestors<TextBoxContent>().Any()) continue;
            if (paragraph.Ancestors<Table>().Any()) continue;

            if (HasBottomBorder(paragraph)) rule = true;

            var pieces = ReadRuns(paragraph, inventory);
            if (pieces.Count == 0) continue;

            var column = columns[ColumnOf(paragraph)];
            column.AddRange(OpeningALine(pieces, column.Count > 0));
        }

        foreach (var drawing in root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
        {
            // O fallback VML repete o mesmo conteúdo; percorrer os dois
            // duplicaria o cabeçalho inteiro.
            if (drawing.Ancestors<AlternateContentFallback>().Any()) continue;
            if (drawing.Ancestors<Table>().Any()) continue;

            // Desenho ancorado tem posição de verdade e pode vir girado: sai
            // como objeto, e não como peça de uma das três colunas. Sem isto, a
            // marca lateral do corpus — 28,6 mm em pé — era achatada numa faixa
            // de 10 mm e desenhada deitada.
            if (AnchorReader.AnchorOf(drawing) is { } anchor)
            {
                foreach (var item in ReadAnchoredDrawing(drawing, owner, anchor, inventory)) floats.Add(item);
                continue;
            }

            ReadDrawing(drawing, owner, columns, ref rule, inventory, contentWidthEmus);
        }

        return new BandDto(columns[0], columns[1], columns[2], rule, floats, rows);
    }

    // --- a grade ------------------------------------------------------------

    /// <summary>
    /// As tabelas do cabeçalho, já com mesclagem e bordas resolvidas.
    /// </summary>
    /// <remarks>
    /// A mesclagem vertical do OOXML não é um `rowSpan`: a célula de cima diz
    /// `restart` e as de baixo aparecem como células vazias com `w:vMerge`. Quem
    /// as desenhasse como células de verdade abriria uma linha vazia debaixo do
    /// logotipo por cada linha mesclada. Aqui elas somem, e a de cima cresce —
    /// que é o que o HTML entende.
    /// </remarks>
    private static List<BandRowDto> ReadGrid(
        OpenXmlPartRootElement root,
        OpenXmlPart owner,
        Inventory inventory)
    {
        var rows = new List<BandRowDto>();

        foreach (var table in root.Elements<Table>())
        {
            var borders = table.GetFirstChild<TableProperties>()?.TableBorders;
            var total = ColumnWidths(table).Sum();
            if (total <= 0) total = 1;

            // A célula aberta em cada coluna da grade, para a continuação da
            // mesclagem saber a quem somar a altura.
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

                    // `w:vMerge` sem `w:val` — ou com `continue` — é a célula de
                    // baixo de uma mesclagem: ela não existe no HTML.
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
                        ReadCellPieces(cell, owner, inventory),
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

    /// <summary>Troca uma célula já montada pela versão que cresceu.</summary>
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
    /// Que lados desta célula têm risco, resolvidos de uma vez.
    /// </summary>
    /// <remarks>
    /// Três origens por lado, da mais forte para a mais fraca: a borda da
    /// própria célula, a borda externa da tabela quando o lado é externo, e a
    /// borda interna quando não é. `nil` na célula apaga o risco que a tabela
    /// pediu — é assim que este cabeçalho junta duas linhas numa só moldura.
    /// </remarks>
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
        // A borda direita usa a externa: só a última coluna a desenha de fato, e
        // sem saber quantas colunas a grade tem preferir a externa erra para o
        // lado de desenhar demais, não de menos.
        if (Drawn(own?.RightBorder, table?.RightBorder)) sides += "r";

        return sides;
    }

    private static bool Drawn(BorderType? own, BorderType? inherited)
    {
        var border = own ?? inherited;
        if (border?.Val?.Value is not { } style) return false;
        return style != BorderValues.None && style != BorderValues.Nil;
    }

    /// <summary>O que está escrito numa célula: imagens e texto, nessa ordem.</summary>
    private static List<PieceDto> ReadCellPieces(
        TableCell cell,
        OpenXmlPart owner,
        Inventory inventory)
    {
        var pieces = new List<PieceDto>();

        foreach (var paragraph in cell.Descendants<Paragraph>())
        {
            foreach (var picture in paragraph.Descendants<Drawing.Pictures.Picture>())
            {
                if (ImagePieceOf(picture, owner) is { } image) pieces.Add(image);
            }

            var run = ReadRuns(paragraph, inventory);
            if (run.Count > 0) pieces.AddRange(OpeningALine(run, pieces.Count > 0));
        }

        return pieces;
    }

    /// <summary>
    /// As peças de um parágrafo, com a primeira marcada como início de linha.
    /// </summary>
    /// <remarks>
    /// Só quando já há alguma coisa antes: a primeira linha da coluna não abre
    /// linha, ela já está numa.
    /// </remarks>
    private static IEnumerable<PieceDto> OpeningALine(List<PieceDto> pieces, bool after)
    {
        for (var index = 0; index < pieces.Count; index++)
        {
            yield return index == 0 && after ? pieces[index] with { Line = true } : pieces[index];
        }
    }

    /// <summary>Uma imagem de dentro da grade, no tamanho que o arquivo pede.</summary>
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
    /// As peças de um desenho ancorado, cada uma na caixa dela.
    /// </summary>
    /// <remarks>
    /// Um cabeçalho corporativo costuma ser um **grupo de formas**: o logotipo,
    /// a caixa do título, a do número da página. Antes daqui todas recebiam a
    /// caixa do grupo inteiro — o logotipo de 48 × 10,5 mm era esticado para os
    /// 177 × 17 mm da faixa toda — e as caixas de texto não saíam, porque só as
    /// imagens eram procuradas: o título do documento sumia do cabeçalho.
    /// </remarks>
    private static IEnumerable<FloatDto> ReadAnchoredDrawing(
        OpenXmlElement drawing,
        OpenXmlPart owner,
        WordDrawing.Anchor anchor,
        Inventory inventory)
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
            foreach (var box in piece.Shape.Descendants<TextBoxContent>())
            {
                foreach (var paragraph in box.Descendants<Paragraph>())
                {
                    var pieces = ReadRuns(paragraph, inventory);
                    if (pieces.Count == 0) continue;

                    var node = Node.Of("paragraph");
                    node.Content = pieces.Select(NodeOf).ToList();
                    content.Add(node);
                }
            }

            if (content.Count > 0) yield return AnchorReader.Describe(anchor, "text", null, content, piece);
        }

        // Uma peça sem `a:xfrm` não tem caixa própria: o desenho é ela sozinha.
        if (AnchorReader.PiecesOf(anchor).Count > 0) yield break;

        foreach (var picture in drawing.Descendants<Drawing.Pictures.Picture>())
        {
            if (ImageSourceOf(picture, owner) is { } src)
            {
                yield return AnchorReader.Describe(anchor, "image", src, null);
            }
        }
    }

    /// <summary>Um pedaço de faixa vira nó de texto, para a caixa desenhá-lo.</summary>
    private static Node NodeOf(PieceDto piece)
    {
        var text = piece.Kind switch
        {
            PieceDto.KindPageNumber => "{n}",
            PieceDto.KindTotalPages => "{total}",
            _ => piece.Text ?? string.Empty,
        };

        var marks = new List<Mark>();
        if (piece.Bold) marks.Add(Mark.Of("bold"));
        if (piece.Italic) marks.Add(Mark.Of("italic"));

        var attributes = new Dictionary<string, System.Text.Json.Nodes.JsonNode?>();
        if (piece.Color is not null) attributes["color"] = piece.Color;
        if (piece.FontSize is not null) attributes["fontSize"] = piece.FontSize;
        if (attributes.Count > 0) marks.Add(new Mark { Type = "textStyle", Attrs = attributes });

        return new Node
        {
            Type = "text",
            Text = text,
            Marks = marks.Count == 0 ? null : marks,
        };
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
        double contentWidthEmus)
    {
        // A posição real na página vem da **âncora**, quando existe. Sem ela, a
        // única coordenada disponível é o `a:off` de dentro do desenho, que num
        // desenho de peça única é sempre zero — e o logotipo do cabeçalho, que
        // no arquivo está a 126,6 mm do começo da coluna, caía no terço do meio
        // por essa conta, quando o Word e o LibreOffice o desenham à direita.
        var anchor = drawing.Descendants<WordDrawing.Anchor>().FirstOrDefault();
        var horizontal = anchor?.GetFirstChild<WordDrawing.HorizontalPosition>();
        var anchorOffset = long.TryParse(horizontal?.PositionOffset?.Text, out var emus)
            ? (double?)emus
            : null;
        var anchorAlign = horizontal?.HorizontalAlignment?.Text;

        var totalWidth = (double?)drawing.Descendants<WordDrawing.Extent>().FirstOrDefault()?.Cx?.Value;
        if (totalWidth is null or <= 0) totalWidth = 1;

        // Coordenadas das peças dentro do grupo, quando há grupo.
        var groupExtent = drawing.Descendants<Drawing.ChildExtents>().FirstOrDefault();
        var span = (double?)groupExtent?.Cx?.Value ?? totalWidth.Value;
        var origin = (double?)drawing.Descendants<Drawing.ChildOffset>().FirstOrDefault()?.X?.Value ?? 0;

        foreach (var shape in drawing.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape>())
        {
            var (offset, width, height) = GeometryOf(shape.Descendants<Drawing.Transform2D>().FirstOrDefault());
            var pieces = shape.Descendants<TextBoxContent>()
                .SelectMany(box => box.Descendants<Paragraph>())
                .SelectMany(paragraph => ReadRuns(paragraph, inventory))
                .ToList();

            if (pieces.Count == 0)
            {
                // Forma vazia, larga e baixa é o filete sob o cabeçalho.
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

    /// <summary>
    /// Em que terço da largura o centro da peça cai.
    /// </summary>
    private static int ColumnFor(
        double offset,
        double width,
        double origin,
        double span,
        double? anchorOffset,
        string? anchorAlign,
        double contentWidthEmus)
    {
        // Alinhamento declarado não precisa de conta nenhuma: o arquivo já diz
        // em que terço a peça está.
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
            // A coordenada de dentro do desenho entra como está. Num desenho de
            // peça única — o caso do logotipo — ela é zero, e a âncora responde
            // sozinha. Num grupo ela está no espaço do grupo e não em EMU da
            // página, o que torna a soma uma aproximação: erra dentro do próprio
            // grupo, nunca sobre em que terço da página o grupo está.
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

    // --- runs ---------------------------------------------------------------

    /// <summary>
    /// Onde estamos dentro de um campo. Entre `separate` e `end` está o último
    /// valor calculado, em cache: copiá-lo junto com o nosso marcador faria o
    /// cabeçalho virar "{n}5" — o marcador mais o número da página em que o
    /// arquivo foi salvo pela última vez.
    /// </summary>
    private sealed class FieldState
    {
        public bool InCachedResult;
    }

    private static List<PieceDto> ReadRuns(Paragraph paragraph, Inventory inventory)
    {
        var pieces = new List<PieceDto>();
        var field = new FieldState();
        Collect(paragraph, pieces, field, inventory);

        // Junta textos vizinhos de mesmo estilo: o Word pica uma frase em vários
        // runs, e sem isto cada pedaço viraria um elemento solto.
        var merged = new List<PieceDto>();
        foreach (var piece in pieces)
        {
            var previous = merged.Count > 0 ? merged[^1] : null;
            if (piece.Kind == PieceDto.KindText && previous is { Kind: PieceDto.KindText } &&
                previous.Bold == piece.Bold && previous.Italic == piece.Italic &&
                previous.Color == piece.Color && previous.FontSize == piece.FontSize)
            {
                merged[^1] = previous with { Text = previous.Text + piece.Text };
                continue;
            }

            merged.Add(piece);
        }

        return merged
            .Where(piece => piece.Kind != PieceDto.KindText || !string.IsNullOrWhiteSpace(piece.Text))
            .ToList();
    }

    private static void Collect(
        OpenXmlElement parent,
        List<PieceDto> pieces,
        FieldState field,
        Inventory inventory)
    {
        foreach (var element in parent.ChildElements)
        {
            switch (element)
            {
                // Desenhos têm passe próprio, que sabe a posição de cada peça.
                // Descer neles aqui traria o mesmo conteúdo uma segunda vez,
                // sem posição — e ele acabaria todo na coluna da esquerda.
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

                case FieldCode code:
                    if (code.Text.Contains("NUMPAGES", StringComparison.Ordinal))
                    {
                        pieces.Add(new PieceDto(PieceDto.KindTotalPages));
                    }
                    else if (code.Text.Contains("PAGE", StringComparison.Ordinal))
                    {
                        pieces.Add(new PieceDto(PieceDto.KindPageNumber));
                    }
                    else
                    {
                        inventory.NoteInvisible(Inventory.HeaderFields);
                    }

                    break;

                case Run run:
                    Collect(run, pieces, field, StyleOf(run.RunProperties), inventory);
                    break;

                default:
                    Collect(element, pieces, field, inventory);
                    break;
            }
        }
    }

    private static void Collect(
        Run run,
        List<PieceDto> pieces,
        FieldState field,
        PieceDto style,
        Inventory inventory)
    {
        foreach (var element in run.ChildElements)
        {
            switch (element)
            {
                case Text text when !field.InCachedResult:
                    pieces.Add(style with { Text = text.Text });
                    break;

                case TabChar when !field.InCachedResult:
                    pieces.Add(style with { Text = " " });
                    break;

                case FieldChar or FieldCode:
                    Collect(run, pieces, field, inventory);
                    return;

                case RunProperties:
                case Text:
                case TabChar:
                    break;
            }
        }
    }

    private static PieceDto StyleOf(RunProperties? properties) => new(
        PieceDto.KindText,
        Bold: RunReader.IsOn(properties?.Bold),
        Italic: RunReader.IsOn(properties?.Italic),
        Color: ColorOf(properties?.Color?.Val?.Value),
        FontSize: SizeOf(properties?.FontSize?.Val?.Value));

    private static string? ColorOf(string? value) =>
        string.IsNullOrWhiteSpace(value) || value.Equals("auto", StringComparison.OrdinalIgnoreCase)
            ? null
            : "#" + value.TrimStart('#').ToLowerInvariant();

    /// <summary>`w:sz` vem em meios-pontos: 40 significa 20 pt.</summary>
    private static string? SizeOf(string? halfPoints) =>
        double.TryParse(halfPoints, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? (value / 2).ToString("0.#", CultureInfo.InvariantCulture) + "pt"
            : null;
}

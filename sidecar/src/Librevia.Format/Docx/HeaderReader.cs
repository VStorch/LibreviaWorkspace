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

    public static BandDto Read(SectionProperties section, MainDocumentPart part, Inventory inventory) =>
        ReadReferenced(section.Elements<HeaderReference>().Select(r => r.Id?.Value), part, inventory);

    public static BandDto ReadFooter(SectionProperties section, MainDocumentPart part, Inventory inventory) =>
        ReadReferenced(section.Elements<FooterReference>().Select(r => r.Id?.Value), part, inventory);

    private static BandDto ReadReferenced(
        IEnumerable<string?> relationshipIds,
        MainDocumentPart part,
        Inventory inventory)
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

            var band = Build(root, owner, inventory);
            // A primeira página costuma ter cabeçalho vazio; vale o primeiro
            // que tenha conteúdo.
            if (!band.IsEmpty) return band;
        }

        return BandDto.Empty();
    }

    private static BandDto Build(OpenXmlPartRootElement root, OpenXmlPart owner, Inventory inventory)
    {
        var columns = new List<PieceDto>[3];
        for (var i = 0; i < 3; i++) columns[i] = [];

        var rule = false;

        foreach (var paragraph in root.Descendants<Paragraph>())
        {
            // Parágrafos de dentro de caixa de texto são tratados junto com a
            // forma que os contém, para herdar a posição dela.
            if (paragraph.Ancestors<TextBoxContent>().Any()) continue;

            if (HasBottomBorder(paragraph)) rule = true;

            var pieces = ReadRuns(paragraph, inventory);
            if (pieces.Count > 0) columns[ColumnOf(paragraph)].AddRange(pieces);
        }

        foreach (var drawing in root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
        {
            // O fallback VML repete o mesmo conteúdo; percorrer os dois
            // duplicaria o cabeçalho inteiro.
            if (drawing.Ancestors<AlternateContentFallback>().Any()) continue;

            ReadDrawing(drawing, owner, columns, ref rule, inventory);
        }

        return new BandDto(columns[0], columns[1], columns[2], rule);
    }

    private static void ReadDrawing(
        OpenXmlElement drawing,
        OpenXmlPart owner,
        List<PieceDto>[] columns,
        ref bool rule,
        Inventory inventory)
    {
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

            columns[ColumnFor(offset, width, origin, span)].AddRange(pieces);
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

            columns[ColumnFor(offset, width, origin, span)].Add(PieceDto.Image(
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
    private static int ColumnFor(double offset, double width, double origin, double span)
    {
        if (span <= 0) return 0;
        var center = (offset - origin + width / 2) / span;
        return center < 1.0 / 3 ? 0 : center < 2.0 / 3 ? 1 : 2;
    }

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
                        inventory.NoteInvisible("campos calculados no cabeçalho");
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

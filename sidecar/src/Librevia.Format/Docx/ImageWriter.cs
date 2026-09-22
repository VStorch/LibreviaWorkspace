using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Imagem do editor → `w:drawing` no fluxo do texto.
/// </summary>
/// <remarks>
/// Saiu de <see cref="ParagraphWriter"/> porque é a única responsabilidade dele
/// que tem contas próprias: decodificar o data URI, escolher o tipo de parte,
/// medir o cabeçalho dos bytes e converter pixel em EMU. Trinta linhas de árvore
/// de DrawingML no meio do escritor de parágrafo escondiam as duas coisas.
/// </remarks>
internal sealed class ImageWriter(MainDocumentPart part, Inventory inventory, int usableWidthPx)
{
    /// <summary>A coluna de uma A4 retrato com margens de uma polegada.</summary>
    internal const int DefaultWidthPx = 624;

    /// <summary>
    /// O próximo `wp:docPr/@id`, contado a partir do maior que o documento já
    /// usa. Zero quer dizer "ainda não olhei o documento".
    /// </summary>
    /// <remarks>
    /// De instância, e não `static`: o servidor atende várias gravações no mesmo
    /// processo, e um contador compartilhado fazia o id da imagem depender de
    /// quantas requisições já tinham passado — e o `++` sobre um campo estático
    /// nem é atômico. Cada gravação tem o seu ImageWriter, e o id volta a ser
    /// função do documento.
    ///
    /// Começa acima do maior `wp:docPr/@id` de `word/document.xml` porque o
    /// desenho ancorado que a gravação preserva leva o id dele junto: repeti-lo é
    /// o que o Word mostra como documento danificado.
    /// </remarks>
    private uint _nextDrawingId;

    public Run? Write(Node node)
    {
        var source = Attr.String(node, "src");
        if (source is null || !source.StartsWith("data:", StringComparison.Ordinal))
        {
            inventory.NoteLoss("imagem sem conteúdo embutido");
            return null;
        }

        // Sem a vírgula não há onde o cabeçalho termina, e o recorte estourava o
        // fim da string: um `src="data:image/png"` truncado derrubava o save
        // inteiro em vez de custar só a imagem.
        var comma = source.IndexOf(',', StringComparison.Ordinal);
        if (comma < 0)
        {
            inventory.NoteLoss("imagem com endereço inválido");
            return null;
        }

        var header = source[5..comma];
        if (!header.EndsWith(";base64", StringComparison.OrdinalIgnoreCase))
        {
            inventory.NoteLoss("imagem em formato não suportado");
            return null;
        }

        var contentType = header[..^";base64".Length];
        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(source[(comma + 1)..]);
        }
        catch (FormatException)
        {
            inventory.NoteLoss("imagem com conteúdo ilegível");
            return null;
        }

        // No OpenXml 3.x `ImagePartType` é classe estática de `PartTypeInfo`,
        // e não mais um enum.
        PartTypeInfo imageType;
        switch (contentType)
        {
            case "image/png": imageType = ImagePartType.Png; break;
            case "image/jpeg": imageType = ImagePartType.Jpeg; break;
            case "image/gif": imageType = ImagePartType.Gif; break;
            case "image/bmp": imageType = ImagePartType.Bmp; break;
            default:
                inventory.NoteLoss($"imagem em {contentType}");
                return null;
        }

        var imagePart = part.AddImagePart(imageType);
        using (var stream = new MemoryStream(bytes))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = part.GetIdOfPart(imagePart);

        var (widthPx, heightPx) = Dimensions(node, bytes);

        // Pixels CSS → EMU: 914400 por polegada, 96 px por polegada.
        var cx = (long)widthPx * 914400 / 96;
        var cy = (long)heightPx * 914400 / 96;

        var id = NextDrawingId();

        // O texto alternativo, que é o que um leitor de tela lê no lugar da
        // imagem. O `descr` só é escrito quando há algo escrito: `descr=""` em
        // toda imagem seria ruído no arquivo e diferença no modelo.
        var description = Attr.String(node, "alt");

        return new Run(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
            new WordDrawing.Inline(
                new WordDrawing.Extent { Cx = cx, Cy = cy },
                new WordDrawing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new WordDrawing.DocProperties
                {
                    Id = id,
                    Name = "Imagem " + id,
                    Description = string.IsNullOrEmpty(description) ? null : description,
                },
                new Drawing.Graphic(
                    new Drawing.GraphicData(
                        new Pictures.Picture(
                            new Pictures.NonVisualPictureProperties(
                                new Pictures.NonVisualDrawingProperties { Id = 0U, Name = "Imagem " + id },
                                new Pictures.NonVisualPictureDrawingProperties()),
                            new Pictures.BlipFill(
                                new Drawing.Blip { Embed = relationshipId },
                                new Drawing.Stretch(new Drawing.FillRectangle())),
                            new Pictures.ShapeProperties(
                                new Drawing.Transform2D(
                                    new Drawing.Offset { X = 0L, Y = 0L },
                                    new Drawing.Extents { Cx = cx, Cy = cy }),
                                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                                {
                                    Preset = Drawing.ShapeTypeValues.Rectangle,
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
            }));
    }

    /// <summary>
    /// Os desenhos do parágrafo original que o leitor entregou ao editor como
    /// imagem do parágrafo, na ordem do arquivo.
    /// </summary>
    /// <remarks>
    /// São os `w:drawing` filhos diretos de um `w:r` — o que está numa caixa de
    /// texto é conteúdo da caixa — que trazem uma imagem e correm com o texto: o
    /// `wp:inline` e o ancorado que o fluxo poria no mesmo lugar. O ancorado com
    /// posição de verdade não entra, porque ele vai para os objetos do
    /// parágrafo e é copiado inteiro por outro caminho.
    /// </remarks>
    public static List<DocumentFormat.OpenXml.Wordprocessing.Drawing> FlowingImagesOf(OpenXmlElement? original)
    {
        if (original is null) return [];

        return [.. original.Descendants<Run>()
            .Where(run => !run.Ancestors<TextBoxContent>().Any())
            .SelectMany(run => run.Elements<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
            .Where(drawing => drawing.Descendants<Drawing.Blip>().Any())
            .Where(drawing => AnchorReader.AnchorOf(drawing) is not { } anchor || AnchorReader.FlowsWithText(anchor))];
    }

    /// <summary>
    /// A imagem que já estava no arquivo volta com o desenho dela, e não com um
    /// desenho novo.
    /// </summary>
    /// <remarks>
    /// Redimensionar uma imagem lida do `.docx` a regravava como se fosse nova:
    /// outro `wp:docPr` — "Imagem 2" no lugar do nome e do id que o documento
    /// dava —, outra parte de imagem com outro relacionamento, e tudo o que este
    /// escritor não sabe gerar ia embora sem aviso: recorte, efeito, borda,
    /// posição do ancorado. O que a pessoa muda numa imagem é o tamanho e o texto
    /// alternativo, e é só isso que muda aqui; o resto é o XML original.
    ///
    /// O par é achado pelo conteúdo, e não pela posição: a imagem que a pessoa
    /// apagou não pode emprestar o desenho dela para a vizinha. Cada desenho
    /// serve a uma imagem só — o que foi usado sai da lista.
    /// </remarks>
    /// <returns>O `w:r` com o desenho original ajustado, ou <c>null</c> quando nenhum confere.</returns>
    public Run? Reuse(Node node, List<DocumentFormat.OpenXml.Wordprocessing.Drawing> candidates)
    {
        var source = Attr.String(node, "src");
        if (source is null) return null;

        var index = candidates.FindIndex(drawing => SourceOf(drawing) == source);
        if (index < 0) return null;

        var original = candidates[index];
        candidates.RemoveAt(index);

        var drawing = (DocumentFormat.OpenXml.Wordprocessing.Drawing)original.CloneNode(true);
        Resize(drawing, Attr.Int(node, "width"), Attr.Int(node, "height"));
        Describe(drawing, Attr.String(node, "alt"));

        // Só a formatação do run vai junto: o texto que dividia o `w:r` com a
        // imagem já chegou ao editor como texto, e é de lá que ele volta.
        var run = new Run();
        if (original.Parent is Run { RunProperties: { } properties })
        {
            run.AppendChild(properties.CloneNode(true));
        }

        run.AppendChild(drawing);
        return run;
    }

    /// <summary>O data URI da imagem do desenho, na mesma forma que o leitor o monta.</summary>
    private string? SourceOf(OpenXmlElement drawing)
    {
        var relationshipId = drawing.Descendants<Drawing.Blip>().FirstOrDefault()?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;

        if (!part.TryGetPartById(relationshipId, out var found) || found is not ImagePart image) return null;

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}";
    }

    /// <summary>
    /// O tamanho novo, no `wp:extent` e no `a:ext` da figura — e só quando mudou.
    /// </summary>
    /// <remarks>
    /// Comparado em pixels, que é a unidade do editor: reconverter a medida que
    /// ninguém tocou mudaria o EMU original por arredondamento, e o arquivo
    /// diferiria sem ninguém ter pedido. No quarto de volta o leitor trocou
    /// largura e altura, e a conta volta a trocá-las.
    /// </remarks>
    private void Resize(OpenXmlElement drawing, int? width, int? height)
    {
        if (width is not > 0 || height is not > 0) return;

        var extent = drawing.Descendants<WordDrawing.Extent>().FirstOrDefault();
        if (extent?.Cx?.Value is not { } cx || extent.Cy?.Value is not { } cy) return;

        var turned = BodyReader.IsQuarterTurned(drawing);
        var (across, down) = turned ? (cy, cx) : (cx, cy);
        if (ToPx(across) == width && ToPx(down) == height) return;

        var (wide, tall) = (width.Value, height.Value);
        if (wide > usableWidthPx)
        {
            tall = Math.Max(1, (int)Math.Round((double)tall * usableWidthPx / wide));
            wide = usableWidthPx;
        }

        var (newCx, newCy) = turned ? (ToEmu(tall), ToEmu(wide)) : (ToEmu(wide), ToEmu(tall));
        extent.Cx = newCx;
        extent.Cy = newCy;

        if (drawing.Descendants<Pictures.ShapeProperties>().FirstOrDefault()?.Transform2D?.Extents is { } extents)
        {
            extents.Cx = newCx;
            extents.Cy = newCy;
        }
    }

    /// <summary>
    /// O texto alternativo no `wp:docPr` — e no `pic:cNvPr` quando ele também o traz.
    /// </summary>
    /// <remarks>
    /// O leitor só entrega o `alt` quando há algo escrito, então a ausência dele
    /// diante de um `descr` preenchido é a pessoa que apagou o texto.
    /// </remarks>
    private static void Describe(OpenXmlElement drawing, string? alt)
    {
        var wanted = string.IsNullOrEmpty(alt) ? null : alt;

        if (drawing.Descendants<WordDrawing.DocProperties>().FirstOrDefault() is { } properties &&
            (properties.Description?.Value is { Length: > 0 } || wanted is not null))
        {
            properties.Description = wanted;
        }

        if (drawing.Descendants<Pictures.NonVisualDrawingProperties>().FirstOrDefault() is { } picture &&
            (picture.Description?.Value is { Length: > 0 } || (wanted is not null && picture.Description is not null)))
        {
            picture.Description = wanted;
        }
    }

    private static int ToPx(long emu) => (int)Math.Round(emu * 96.0 / 914400);

    private static long ToEmu(int px) => (long)px * 914400 / 96;

    private uint NextDrawingId()
    {
        if (_nextDrawingId == 0)
        {
            _nextDrawingId = (part.Document?.Descendants<WordDrawing.DocProperties>()
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max() ?? 0U) + 1;
        }

        return _nextDrawingId++;
    }

    /// <summary>
    /// De que tamanho a imagem entra no arquivo.
    /// </summary>
    /// <remarks>
    /// A imagem que vem do `.docx` traz as duas medidas, porque o leitor as lê do
    /// `wp:extent`: é o tamanho que o documento pede, que não precisa ser o do
    /// arquivo. A que a pessoa insere pela barra de ferramentas não traz nenhuma,
    /// e aí valem as do cabeçalho dos próprios bytes — antes daqui o escritor
    /// chutava 600 × 450, e uma captura quadrada saía deitada.
    ///
    /// O teto é a largura da coluna de texto. Uma captura de tela de 1920 px
    /// entraria com 50 cm de largura e o Word a desenharia estourando as duas
    /// margens; encolhida na proporção, ela cabe onde a pessoa a viu caber.
    /// </remarks>
    private (int Width, int Height) Dimensions(Node node, byte[] bytes)
    {
        var measured = ImageSize.Of(bytes);

        var width = Attr.Int(node, "width") ?? measured?.Width ?? DefaultWidthPx;
        var height = Attr.Int(node, "height")
                     ?? (measured is { } size && size.Width > 0
                         ? (int)Math.Round((double)width * size.Height / size.Width)
                         : (int)Math.Round(width * 0.75));

        if (width <= 0 || height <= 0) return (DefaultWidthPx, (int)(DefaultWidthPx * 0.75));
        if (width <= usableWidthPx) return (width, height);

        return (usableWidthPx, Math.Max(1, (int)Math.Round((double)height * usableWidthPx / width)));
    }
}

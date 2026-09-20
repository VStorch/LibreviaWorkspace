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

        return new Run(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
            new WordDrawing.Inline(
                new WordDrawing.Extent { Cx = cx, Cy = cy },
                new WordDrawing.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new WordDrawing.DocProperties { Id = id, Name = "Imagem " + id },
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

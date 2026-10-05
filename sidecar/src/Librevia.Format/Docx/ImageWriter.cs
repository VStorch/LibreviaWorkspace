using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Editor image → a flowing <c>w:drawing</c>: decodes the data URI, picks the part type, measures
/// the bytes and converts to EMU.
/// </summary>
/// <param name="owner">The document, or the notes part.</param>
internal sealed class ImageWriter(MainDocumentPart part, Inventory inventory, int usableWidthPx, OpenXmlPart? owner = null)
{
    /// <summary>A4 portrait with one-inch margins.</summary>
    internal const int DefaultWidthPx = 624;

    /// <summary>Without a height on the node or in the bytes, a 4:3 ratio.</summary>
    private const double DefaultHeightPerWidth = 0.75;

    /// <summary>
    /// Above the document's highest <c>wp:docPr/@id</c>: a preserved anchored object keeps its own,
    /// and Word shows a repeated one as a damaged document. Per instance, because the server
    /// handles several saves in the same process.
    /// </summary>
    private uint _nextDrawingId;

    public Run? Write(Node node)
    {
        var source = Attr.String(node, "src");
        if (source is null || !source.StartsWith("data:", StringComparison.Ordinal))
        {
            inventory.NoteLoss("imagem sem conteúdo embutido");
            return null;
        }

        // Without the comma, slicing would overflow the string: a truncated image only costs the
        // image.
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

        // In OpenXml 3.x `ImagePartType` is a class, not an enum.
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

        var imagePart = (owner ?? part) switch
        {
            FootnotesPart footnotes => footnotes.AddImagePart(imageType),
            EndnotesPart endnotes => endnotes.AddImagePart(imageType),
            _ => part.AddImagePart(imageType),
        };
        using (var stream = new MemoryStream(bytes))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = (owner ?? part).GetIdOfPart(imagePart);

        var (widthPx, heightPx) = Dimensions(node, bytes);

        var cx = widthPx * Unit.EmusPerPixel;
        var cy = heightPx * Unit.EmusPerPixel;

        var id = NextDrawingId();

        // `descr` only when something is written.
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
    /// Direct <c>w:drawing</c> children of <c>w:r</c> with an image flowing with the text, in file
    /// order; anchored ones with a position are copied another way.
    /// </summary>
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
    /// A file image goes back with its drawing: crop, effect, border and name stay, and only size
    /// and alt text change. The pair is found by content, and each drawing serves a single image.
    /// </summary>
    /// <returns>The <c>w:r</c> with the adjusted drawing, or <c>null</c> when none
    /// matches.</returns>
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

        // Only the run formatting goes along: the neighbouring text comes back through the editor.
        var run = new Run();
        if (original.Parent is Run { RunProperties: { } properties })
        {
            run.AppendChild(properties.CloneNode(true));
        }

        run.AppendChild(drawing);
        return run;
    }

    /// <summary>In the same shape the reader builds.</summary>
    private string? SourceOf(OpenXmlElement drawing)
    {
        var relationshipId = drawing.Descendants<Drawing.Blip>().FirstOrDefault()?.Embed?.Value;
        if (string.IsNullOrEmpty(relationshipId)) return null;

        if (!(owner ?? part).TryGetPartById(relationshipId, out var found) || found is not ImagePart image) return null;

        using var stream = image.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return $"data:{image.ContentType};base64,{Convert.ToBase64String(buffer.ToArray())}";
    }

    /// <summary>
    /// Only when it changed, compared in pixels: reconverting what nobody touched would change the
    /// EMU. On a quarter turn, width and height swap back.
    /// </summary>
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
    /// An absent <c>alt</c> against a filled <c>descr</c> is the user having cleared it.
    /// </summary>
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

    private static int ToPx(long emu) => (int)Math.Round(Unit.EmusToPixels(emu));

    private static long ToEmu(int px) => px * Unit.EmusPerPixel;

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
    /// A <c>.docx</c> image carries the <c>wp:extent</c> measures; an inserted one, those of the
    /// byte header. The ceiling is the text column, keeping the ratio.
    /// </summary>
    private (int Width, int Height) Dimensions(Node node, byte[] bytes)
    {
        var measured = ImageSize.Of(bytes);

        var width = Attr.Int(node, "width") ?? measured?.Width ?? DefaultWidthPx;
        var height = Attr.Int(node, "height")
                     ?? (measured is { } size && size.Width > 0
                         ? (int)Math.Round((double)width * size.Height / size.Width)
                         : (int)Math.Round(width * DefaultHeightPerWidth));

        if (width <= 0 || height <= 0) return (DefaultWidthPx, (int)(DefaultWidthPx * DefaultHeightPerWidth));
        if (width <= usableWidthPx) return (width, height);

        return (usableWidthPx, Math.Max(1, (int)Math.Round((double)height * usableWidthPx / width)));
    }
}

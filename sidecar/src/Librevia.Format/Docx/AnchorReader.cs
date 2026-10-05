using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// In millimetres: screen and paper draw at different resolutions, and each converts once.
/// </summary>
public sealed record FloatDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("src")] string? Src,
    [property: JsonPropertyName("content")] List<Node>? Content,
    [property: JsonPropertyName("widthMm")] double WidthMm,
    [property: JsonPropertyName("heightMm")] double HeightMm,
    /// <summary>Degrees, clockwise, as in CSS.</summary>
    [property: JsonPropertyName("rotation")] double Rotation,
    [property: JsonPropertyName("hFrom")] string HorizontalFrom,
    [property: JsonPropertyName("hOffsetMm")] double? HorizontalOffsetMm,
    [property: JsonPropertyName("hAlign")] string? HorizontalAlign,
    [property: JsonPropertyName("vFrom")] string VerticalFrom,
    [property: JsonPropertyName("vOffsetMm")] double? VerticalOffsetMm,
    [property: JsonPropertyName("vAlign")] string? VerticalAlign,
    /// <summary>Cover decoration, watermark.</summary>
    [property: JsonPropertyName("behind")] bool Behind,
    [property: JsonPropertyName("wrap")] string Wrap,
    /// <summary>The piece's position inside the group, added after resolving the anchor.</summary>
    [property: JsonPropertyName("dxMm")] double DxMm = 0,
    [property: JsonPropertyName("dyMm")] double DyMm = 0,
    /// <summary>
    /// Band objects only; the box is regenerated whole, because typing opens and closes paragraphs.
    /// </summary>
    [property: JsonPropertyName("bid")] string? BoxId = null,
    /// <summary>
    /// Solid color and stroke only; the rest goes to the inventory (<see cref="ShapeLook"/>).
    /// </summary>
    [property: JsonPropertyName("fill")] string? Fill = null,
    [property: JsonPropertyName("line")] string? Line = null,
    [property: JsonPropertyName("lineWidthPt")] double LineWidthPt = 0,
    [property: JsonPropertyName("dash")] bool Dash = false);

    /// <summary>On the page ruler.</summary>
public sealed record AnchoredPiece(
    OpenXmlElement Shape,
    double DxEmus,
    double DyEmus,
    double WidthEmus,
    double HeightEmus,
    double Rotation);

/// <c>wp:anchor</c> → origin and offset per axis, unresolved: the most common origin is the
/// paragraph, which only has a position after pagination. Rotation comes out in degrees and does
/// not touch the measures, as in Word and <c>transform: rotate()</c>.
public static class AnchorReader
{
    private const double EmusPerMillimeter = Unit.EmusPerInch / Unit.MillimetersPerInch;

    private const double RotationUnitsPerDegree = 60000;

    private const double FlowToleranceMm = 2;

    public static Anchor.Anchor? AnchorOf(OpenXmlElement drawing) =>
        drawing.Descendants<Anchor.Anchor>().FirstOrDefault();

    /// <c>wp:anchor</c> does not mean "out of the flow": LibreOffice writes an image in its own
    /// paragraph that way. Only an object that does not move with the paragraph, moves away from
    /// it, sits behind the text or uses <c>wrapNone</c> has a real position; the rest is a block.
    public static bool FlowsWithText(Anchor.Anchor anchor)
    {
        if (anchor.BehindDoc?.Value == true) return false;
        if (WrapOf(anchor) == "none") return false;

        var vertical = anchor.GetFirstChild<Anchor.VerticalPosition>();
        var from = vertical?.RelativeFrom?.Value;
        if (from is not null &&
            from != Anchor.VerticalRelativePositionValues.Paragraph &&
            from != Anchor.VerticalRelativePositionValues.Line)
        {
            return false;
        }

        // Vertical alignment is a declared position.
        if (vertical?.VerticalAlignment is not null) return false;
        if (Math.Abs(OffsetMillimeters(vertical?.PositionOffset?.Text) ?? 0) > FlowToleranceMm) return false;

        var horizontal = anchor.GetFirstChild<Anchor.HorizontalPosition>();

        // Aligned in the column is where the paragraph would already put it.
        if (horizontal?.HorizontalAlignment is not null) return true;

        var side = horizontal?.RelativeFrom?.Value;
        if (side is not null &&
            side != Anchor.HorizontalRelativePositionValues.Column &&
            side != Anchor.HorizontalRelativePositionValues.Margin &&
            side != Anchor.HorizontalRelativePositionValues.Character)
        {
            return false;
        }

        return Math.Abs(OffsetMillimeters(horizontal?.PositionOffset?.Text) ?? 0) <= FlowToleranceMm;
    }

    public static FloatDto Describe(
        Anchor.Anchor anchor,
        string kind,
        string? src,
        List<Node>? content,
        AnchoredPiece? piece = null)
    {
        var extent = anchor.Descendants<Anchor.Extent>().FirstOrDefault();
        var horizontal = anchor.GetFirstChild<Anchor.HorizontalPosition>();
        var vertical = anchor.GetFirstChild<Anchor.VerticalPosition>();

        return new FloatDto(
            Kind: kind,
            Src: src,
            Content: content,
            WidthMm: piece is null ? Millimeters(extent?.Cx?.Value) : Round(piece.WidthEmus),
            HeightMm: piece is null ? Millimeters(extent?.Cy?.Value) : Round(piece.HeightEmus),
            Rotation: piece?.Rotation ?? RotationOf(anchor),
            HorizontalFrom: horizontal?.RelativeFrom?.ToString() ?? "column",
            HorizontalOffsetMm: OffsetMillimeters(horizontal?.PositionOffset?.Text),
            HorizontalAlign: horizontal?.HorizontalAlignment?.Text,
            VerticalFrom: vertical?.RelativeFrom?.ToString() ?? "paragraph",
            VerticalOffsetMm: OffsetMillimeters(vertical?.PositionOffset?.Text),
            VerticalAlign: vertical?.VerticalAlignment?.Text,
            Behind: anchor.BehindDoc?.Value ?? false,
            Wrap: WrapOf(anchor),
            DxMm: piece is null ? 0 : Round(piece.DxEmus),
            DyMm: piece is null ? 0 : Round(piece.DyEmus));
    }

    /// <summary>
    /// The anchor says where the group is; <c>a:chOff</c> and <c>a:chExt</c> give the ruler for the
    /// inner coordinates. A single-piece drawing comes out with zero offset.
    /// </summary>
    public static List<AnchoredPiece> PiecesOf(Anchor.Anchor anchor)
    {
        var extent = anchor.Descendants<Anchor.Extent>().FirstOrDefault();
        // Only a group `a:xfrm` carries `a:chOff`/`a:chExt`.
        var group = anchor.Descendants<Drawing.TransformGroup>().FirstOrDefault();

        var (scaleX, scaleY) = (1.0, 1.0);
        var (originX, originY) = (0.0, 0.0);

        if (group?.ChildExtents is { } child)
        {
            originX = group.ChildOffset?.X?.Value ?? 0;
            originY = group.ChildOffset?.Y?.Value ?? 0;
            if (child.Cx?.Value is > 0 && extent?.Cx?.Value is { } across)
            {
                scaleX = across / (double)child.Cx.Value;
            }

            if (child.Cy?.Value is > 0 && extent?.Cy?.Value is { } down)
            {
                scaleY = down / (double)child.Cy.Value;
            }
        }

        var pieces = new List<AnchoredPiece>();
        foreach (var shape in Shapes(anchor))
        {
            var transform = shape.Descendants<Drawing.Transform2D>().FirstOrDefault();
            if (transform?.Extents is not { } size) continue;

            pieces.Add(new AnchoredPiece(
                shape,
                ((transform.Offset?.X?.Value ?? 0) - originX) * scaleX,
                ((transform.Offset?.Y?.Value ?? 0) - originY) * scaleY,
                (size.Cx?.Value ?? 0) * scaleX,
                (size.Cy?.Value ?? 0) * scaleY,
                DegreesOf(transform.Rotation?.Value)));
        }

        return pieces;
    }

    private static IEnumerable<OpenXmlElement> Shapes(Anchor.Anchor anchor) =>
        anchor.Descendants<OpenXmlElement>()
            .Where(element =>
                element is Drawing.Pictures.Picture
                    or DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape);

    /// <summary>
    /// Only the mode name, uninterpreted: whoever draws decides what to reproduce.
    /// </summary>
    internal static string WrapOf(Anchor.Anchor anchor)
    {
        if (anchor.GetFirstChild<Anchor.WrapNone>() is not null) return "none";
        if (anchor.GetFirstChild<Anchor.WrapSquare>() is not null) return "square";
        if (anchor.GetFirstChild<Anchor.WrapTight>() is not null) return "tight";
        if (anchor.GetFirstChild<Anchor.WrapThrough>() is not null) return "through";
        if (anchor.GetFirstChild<Anchor.WrapTopBottom>() is not null) return "topAndBottom";
        return "none";
    }

    private static double RotationOf(OpenXmlElement drawing) =>
        DegreesOf(drawing.Descendants<Drawing.Transform2D>().FirstOrDefault()?.Rotation?.Value);

    private static double DegreesOf(int? rotation)
    {
        if (rotation is null) return 0;

        var degrees = (rotation.Value / RotationUnitsPerDegree % 360 + 360) % 360;
        return Math.Round(degrees, 2);
    }

    private static double Round(double emus) => Math.Round(emus / EmusPerMillimeter, 2);

    private static double Millimeters(long? emus) =>
        emus is null ? 0 : Math.Round(emus.Value / EmusPerMillimeter, 2);

    /// <summary>Negative means the object reaching into the margin.</summary>
    private static double? OffsetMillimeters(string? text) =>
        long.TryParse(text, out var emus) ? Math.Round(emus / EmusPerMillimeter, 2) : null;
}

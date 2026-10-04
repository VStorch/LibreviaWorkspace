using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>Em milímetros: tela e papel desenham em resoluções diferentes, e cada um converte uma vez.</summary>
public sealed record FloatDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("src")] string? Src,
    [property: JsonPropertyName("content")] List<Node>? Content,
    [property: JsonPropertyName("widthMm")] double WidthMm,
    [property: JsonPropertyName("heightMm")] double HeightMm,
    /// <summary>Graus, sentido horário, como o CSS.</summary>
    [property: JsonPropertyName("rotation")] double Rotation,
    [property: JsonPropertyName("hFrom")] string HorizontalFrom,
    [property: JsonPropertyName("hOffsetMm")] double? HorizontalOffsetMm,
    [property: JsonPropertyName("hAlign")] string? HorizontalAlign,
    [property: JsonPropertyName("vFrom")] string VerticalFrom,
    [property: JsonPropertyName("vOffsetMm")] double? VerticalOffsetMm,
    [property: JsonPropertyName("vAlign")] string? VerticalAlign,
    /// <summary>Decoração de capa, marca d'água.</summary>
    [property: JsonPropertyName("behind")] bool Behind,
    [property: JsonPropertyName("wrap")] string Wrap,
    /// <summary>A posição da peça dentro do grupo, somada depois de resolver a âncora.</summary>
    [property: JsonPropertyName("dxMm")] double DxMm = 0,
    [property: JsonPropertyName("dyMm")] double DyMm = 0,
    /// <summary>Só nos objetos de faixa; a caixa é regenerada inteira, porque digitar abre e fecha parágrafos.</summary>
    [property: JsonPropertyName("bid")] string? BoxId = null,
    /// <summary>Só cor e traço sólidos; o resto vai ao inventário (<see cref="ShapeLook"/>).</summary>
    [property: JsonPropertyName("fill")] string? Fill = null,
    [property: JsonPropertyName("line")] string? Line = null,
    [property: JsonPropertyName("lineWidthPt")] double LineWidthPt = 0,
    [property: JsonPropertyName("dash")] bool Dash = false);

    /// <summary>Na régua da página.</summary>
public sealed record AnchoredPiece(
    OpenXmlElement Shape,
    double DxEmus,
    double DyEmus,
    double WidthEmus,
    double HeightEmus,
    double Rotation);

/// <summary>
/// <c>wp:anchor</c> → origem e deslocamento por eixo, sem resolver: a origem mais
/// comum é o parágrafo, que só tem posição depois de paginar. A rotação sai em
/// graus e não mexe nas medidas, como o Word e o <c>transform: rotate()</c>.
/// </summary>
public static class AnchorReader
{
    private const double EmusPerMillimeter = 914400 / 25.4;

    private const double RotationUnitsPerDegree = 60000;

    private const double FlowToleranceMm = 2;

    public static Anchor.Anchor? AnchorOf(OpenXmlElement drawing) =>
        drawing.Descendants<Anchor.Anchor>().FirstOrDefault();

    /// <summary>
    /// <c>wp:anchor</c> não é "fora do fluxo": o LibreOffice grava assim a imagem no
    /// próprio parágrafo. Tem posição de verdade quem não anda com o parágrafo, se
    /// afasta dele, fica atrás do texto ou usa <c>wrapNone</c>; o resto é bloco.
    /// </summary>
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

        // Alinhamento vertical é posição declarada.
        if (vertical?.VerticalAlignment is not null) return false;
        if (Math.Abs(OffsetMillimeters(vertical?.PositionOffset?.Text) ?? 0) > FlowToleranceMm) return false;

        var horizontal = anchor.GetFirstChild<Anchor.HorizontalPosition>();

        // Alinhado na coluna é onde o parágrafo já o poria.
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
    /// A âncora diz onde está o grupo; <c>a:chOff</c> e <c>a:chExt</c> dão a régua das
    /// coordenadas de dentro. O desenho de peça única sai com deslocamento zero.
    /// </summary>
    public static List<AnchoredPiece> PiecesOf(Anchor.Anchor anchor)
    {
        var extent = anchor.Descendants<Anchor.Extent>().FirstOrDefault();
        // Só o `a:xfrm` de grupo carrega `a:chOff`/`a:chExt`.
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

    /// <summary>Só o nome do modo, sem interpretar: quem desenha decide o que reproduz.</summary>
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

    /// <summary>Negativo é o objeto saindo para a margem.</summary>
    private static double? OffsetMillimeters(string? text) =>
        long.TryParse(text, out var emus) ? Math.Round(emus / EmusPerMillimeter, 2) : null;
}

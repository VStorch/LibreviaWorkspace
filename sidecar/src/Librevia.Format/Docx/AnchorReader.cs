using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Um objeto que não está no fluxo do texto: imagem ou caixa ancorada.
/// </summary>
/// <remarks>
/// Medidas em milímetros, e não em pixels, porque quem as consome desenha em
/// dois lugares com resoluções diferentes — a tela e o papel. Converter uma vez
/// aqui deixaria um dos dois arredondando de volta.
/// </remarks>
public sealed record FloatDto(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("src")] string? Src,
    [property: JsonPropertyName("content")] List<Node>? Content,
    [property: JsonPropertyName("widthMm")] double WidthMm,
    [property: JsonPropertyName("heightMm")] double HeightMm,
    /// <summary>Graus, sentido horário, como o CSS espera.</summary>
    [property: JsonPropertyName("rotation")] double Rotation,
    [property: JsonPropertyName("hFrom")] string HorizontalFrom,
    [property: JsonPropertyName("hOffsetMm")] double? HorizontalOffsetMm,
    [property: JsonPropertyName("hAlign")] string? HorizontalAlign,
    [property: JsonPropertyName("vFrom")] string VerticalFrom,
    [property: JsonPropertyName("vOffsetMm")] double? VerticalOffsetMm,
    [property: JsonPropertyName("vAlign")] string? VerticalAlign,
    /// <summary>Atrás do texto: decoração de capa, marca d'água.</summary>
    [property: JsonPropertyName("behind")] bool Behind,
    [property: JsonPropertyName("wrap")] string Wrap);

/// <summary>
/// `wp:anchor` → onde o objeto cai na folha.
/// </summary>
/// <remarks>
/// O OOXML posiciona um objeto ancorado em relação a **quatro** origens
/// possíveis por eixo — a margem, a coluna, a página e o parágrafo — e a origem
/// vertical mais comum é justamente a que só existe depois de paginar. Por isso
/// nada é resolvido aqui: este leitor entrega a origem e o deslocamento, e quem
/// desenha faz a conta, porque só ele sabe em que folha o parágrafo âncora caiu.
///
/// A rotação sai em graus e **não** é aplicada às medidas. O Word posiciona a
/// caixa sem girar e depois a gira em torno do centro, que é exatamente o que
/// `transform: rotate()` faz — mexer nas medidas aqui desalinharia as duas.
/// </remarks>
public static class AnchorReader
{
    private const double EmusPerMillimeter = 914400 / 25.4;

    /// <summary>60000 avos de grau é a unidade de `a:rot`.</summary>
    private const double RotationUnitsPerDegree = 60000;

    /// <summary>O desenho é ancorado, e não uma imagem no meio da linha?</summary>
    public static Anchor.Anchor? AnchorOf(OpenXmlElement drawing) =>
        drawing.Descendants<Anchor.Anchor>().FirstOrDefault();

    public static FloatDto Describe(
        Anchor.Anchor anchor,
        string kind,
        string? src,
        List<Node>? content)
    {
        var extent = anchor.Descendants<Anchor.Extent>().FirstOrDefault();
        var horizontal = anchor.GetFirstChild<Anchor.HorizontalPosition>();
        var vertical = anchor.GetFirstChild<Anchor.VerticalPosition>();

        return new FloatDto(
            Kind: kind,
            Src: src,
            Content: content,
            WidthMm: Millimeters(extent?.Cx?.Value),
            HeightMm: Millimeters(extent?.Cy?.Value),
            Rotation: RotationOf(anchor),
            HorizontalFrom: horizontal?.RelativeFrom?.ToString() ?? "column",
            HorizontalOffsetMm: OffsetMillimeters(horizontal?.PositionOffset?.Text),
            HorizontalAlign: horizontal?.HorizontalAlignment?.Text,
            VerticalFrom: vertical?.RelativeFrom?.ToString() ?? "paragraph",
            VerticalOffsetMm: OffsetMillimeters(vertical?.PositionOffset?.Text),
            VerticalAlign: vertical?.VerticalAlignment?.Text,
            Behind: anchor.BehindDoc?.Value ?? false,
            Wrap: WrapOf(anchor));
    }

    /// <summary>
    /// Como o texto se comporta em volta.
    /// </summary>
    /// <remarks>
    /// Só o nome do modo, sem interpretá-lo: quem desenha decide o que consegue
    /// reproduzir. Hoje todos são desenhados por cima ou por baixo, sem o texto
    /// contornar — dizer o modo mesmo assim deixa a diferença registrada no
    /// modelo, em vez de perdida no leitor.
    /// </remarks>
    private static string WrapOf(Anchor.Anchor anchor)
    {
        if (anchor.GetFirstChild<Anchor.WrapNone>() is not null) return "none";
        if (anchor.GetFirstChild<Anchor.WrapSquare>() is not null) return "square";
        if (anchor.GetFirstChild<Anchor.WrapTight>() is not null) return "tight";
        if (anchor.GetFirstChild<Anchor.WrapThrough>() is not null) return "through";
        if (anchor.GetFirstChild<Anchor.WrapTopBottom>() is not null) return "topAndBottom";
        return "none";
    }

    private static double RotationOf(OpenXmlElement drawing)
    {
        var rotation = drawing.Descendants<Drawing.Transform2D>().FirstOrDefault()?.Rotation?.Value;
        if (rotation is null) return 0;

        var degrees = (rotation.Value / RotationUnitsPerDegree % 360 + 360) % 360;
        return Math.Round(degrees, 2);
    }

    private static double Millimeters(long? emus) =>
        emus is null ? 0 : Math.Round(emus.Value / EmusPerMillimeter, 2);

    /// <summary>
    /// O deslocamento pode ser negativo — é como o objeto sai para a margem.
    /// </summary>
    private static double? OffsetMillimeters(string? text) =>
        long.TryParse(text, out var emus) ? Math.Round(emus / EmusPerMillimeter, 2) : null;
}

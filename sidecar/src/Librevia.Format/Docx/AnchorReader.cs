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
    [property: JsonPropertyName("wrap")] string Wrap,
    /// <summary>
    /// Deslocamento da peça dentro do desenho, somado depois de resolver a
    /// âncora. Vem de um grupo de formas: a âncora posiciona o grupo, e cada
    /// peça tem a sua coordenada dentro dele.
    /// </summary>
    [property: JsonPropertyName("dxMm")] double DxMm = 0,
    [property: JsonPropertyName("dyMm")] double DyMm = 0,
    /// <summary>
    /// Onde a caixa deste objeto mora, quando ela é editável.
    /// </summary>
    /// <remarks>
    /// Só objetos de faixa o trazem: os do corpo voltam pelo parágrafo que os
    /// ancora, que já tem o `oid`. A caixa é regenerada por inteiro quando o
    /// texto muda — digitar abre e fecha parágrafos dentro dela, e endereçar
    /// parágrafo a parágrafo quebraria na primeira tecla Enter.
    /// </remarks>
    [property: JsonPropertyName("bid")] string? BoxId = null);

/// <summary>Uma peça de dentro de um desenho, na régua da página.</summary>
public sealed record AnchoredPiece(
    OpenXmlElement Shape,
    double DxEmus,
    double DyEmus,
    double WidthEmus,
    double HeightEmus,
    double Rotation);

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

    /// <summary>Meio centímetro de folga: posição de verdade está longe disso.</summary>
    private const double FlowToleranceMm = 2;

    /// <summary>O desenho é ancorado, e não uma imagem no meio da linha?</summary>
    public static Anchor.Anchor? AnchorOf(OpenXmlElement drawing) =>
        drawing.Descendants<Anchor.Anchor>().FirstOrDefault();

    /// <summary>
    /// O objeto ancorado está onde o fluxo o poria de qualquer jeito?
    /// </summary>
    /// <remarks>
    /// `wp:anchor` não quer dizer "fora do fluxo". É assim que o LibreOffice
    /// grava **imagem no próprio parágrafo**: ancorada ao parágrafo, sem
    /// deslocamento vertical, centralizada na coluna e com a largura dela. Um
    /// documento de trinta capturas de tela é feito só disso.
    ///
    /// Tratar essas como posição na folha é o pior dos dois mundos: elas deixam
    /// de ocupar altura, o texto se fecha por cima, o documento encolhe de doze
    /// folhas para quatro e as imagens acabam empilhadas umas sobre as outras.
    ///
    /// A pergunta que separa os dois casos não é o modo de contorno sozinho, é
    /// **onde o objeto está**: quem não anda com o parágrafo, quem se afasta
    /// dele, quem fica atrás do texto ou quem deixa o texto passar por baixo
    /// (`wrapNone`) tem posição de verdade. O resto está no lugar em que o
    /// fluxo já o poria, e é como bloco que ele é desenhado certo.
    /// </remarks>
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

        // "No alto da página", "no meio da margem": alinhamento vertical é
        // posição declarada, e não segue o parágrafo.
        if (vertical?.VerticalAlignment is not null) return false;
        if (Math.Abs(OffsetMillimeters(vertical?.PositionOffset?.Text) ?? 0) > FlowToleranceMm) return false;

        var horizontal = anchor.GetFirstChild<Anchor.HorizontalPosition>();

        // Alinhado na coluna é onde o parágrafo já o poria — inclusive
        // centralizado, que é como a imagem de largura inteira é gravada.
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
    /// As peças de dentro de um desenho ancorado, cada uma na sua caixa.
    /// </summary>
    /// <remarks>
    /// Um cabeçalho corporativo costuma ser **um grupo de formas**: o logotipo,
    /// a caixa do título, a do número da página e um par de filetes, cada um com
    /// coordenada própria dentro do grupo. A âncora diz onde o grupo está e que
    /// tamanho ele tem; `a:chOff` e `a:chExt` dizem em que régua as coordenadas
    /// de dentro foram escritas.
    ///
    /// Sem desembrulhar, cada peça recebia a caixa do grupo inteiro: o logotipo
    /// de 48 × 10,5 mm era esticado para os 177 × 17 mm da faixa toda, e as
    /// caixas de texto — o título do documento entre elas — não saíam de lugar
    /// nenhum, porque quem lia só procurava imagens.
    ///
    /// Desenho de peça única cai aqui também, e sai com deslocamento zero: o
    /// `a:off` dele é a origem, e a conta dá a caixa da própria âncora.
    /// </remarks>
    public static List<AnchoredPiece> PiecesOf(Anchor.Anchor anchor)
    {
        var extent = anchor.Descendants<Anchor.Extent>().FirstOrDefault();
        // `a:xfrm` de grupo é outro elemento: só ele carrega `a:chOff`/`a:chExt`,
        // que é a régua em que as coordenadas de dentro foram escritas.
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

    /// <summary>Imagens e caixas de texto, na ordem em que o arquivo as traz.</summary>
    private static IEnumerable<OpenXmlElement> Shapes(Anchor.Anchor anchor) =>
        anchor.Descendants<OpenXmlElement>()
            .Where(element =>
                element is Drawing.Pictures.Picture
                    or DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape);

    /// <summary>
    /// Como o texto se comporta em volta.
    /// </summary>
    /// <remarks>
    /// Só o nome do modo, sem interpretá-lo: quem desenha decide o que consegue
    /// reproduzir. Hoje todos são desenhados por cima ou por baixo, sem o texto
    /// contornar — dizer o modo mesmo assim deixa a diferença registrada no
    /// modelo, em vez de perdida no leitor.
    /// </remarks>
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

    /// <summary>
    /// O deslocamento pode ser negativo — é como o objeto sai para a margem.
    /// </summary>
    private static double? OffsetMillimeters(string? text) =>
        long.TryParse(text, out var emus) ? Math.Round(emus / EmusPerMillimeter, 2) : null;
}

using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

public sealed record PageSetupDto(
    [property: JsonPropertyName("size")] string Size,
    [property: JsonPropertyName("orientation")] string Orientation,
    [property: JsonPropertyName("margins")] MarginsDto Margins,
    [property: JsonPropertyName("headerBand")] BandDto? Header,
    [property: JsonPropertyName("footerBand")] BandDto? Footer,
    // Campos novos, e não uma troca de forma do par acima: um `.sdoc` gravado
    // antes daqui não os tem, e como são opcionais continua abrindo.
    [property: JsonPropertyName("firstHeaderBand")] BandDto? FirstHeader = null,
    [property: JsonPropertyName("firstFooterBand")] BandDto? FirstFooter = null,
    [property: JsonPropertyName("evenHeaderBand")] BandDto? EvenHeader = null,
    [property: JsonPropertyName("evenFooterBand")] BandDto? EvenFooter = null);

public sealed record MarginsDto(
    [property: JsonPropertyName("top")] double Top,
    [property: JsonPropertyName("right")] double Right,
    [property: JsonPropertyName("bottom")] double Bottom,
    [property: JsonPropertyName("left")] double Left);

/// <summary>
/// `w:sectPr` → configuração de página.
/// </summary>
public static class PageReader
{
    private const double TwipsPerMillimeter = 1440 / 25.4;

    /// <summary>1 twip = 1/1440 polegada; 1 polegada = 914400 EMU.</summary>
    private const double EmusPerTwip = 914400.0 / 1440;

    /// <summary>Tolerância ao comparar seções: 1 twip é 0,018 mm.</summary>
    private const int GeometryTolerance = 2;

    public static PageSetupDto Read(Body body, MainDocumentPart part, Inventory inventory)
    {
        var sections = body.Descendants<SectionProperties>().ToList();
        if (sections.Count == 0) return Default();

        // Seções consecutivas com a mesma geometria são artefato do
        // LibreOffice, não intenção do autor: o documento de 15 páginas do
        // corpus tem sete, todas idênticas. Só há perda quando divergem.
        // Ver docs/01-corpus-docx.md, Descoberta 5.
        if (sections.Count > 1 && !AllShareGeometry(sections))
        {
            inventory.NoteLoss("seções com tamanho ou margem diferentes (o documento usará a primeira)");
        }

        var section = sections[0];
        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        var margin = section.GetFirstChild<PageMargin>();

        var landscape = size?.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        var widthTwips = (double?)size?.Width?.Value ?? 11906;
        var heightTwips = (double?)size?.Height?.Value ?? 16838;

        // A faixa decide em que terço cada peça cai comparando a posição dela
        // com a largura da coluna de texto. Sem esta medida, o logotipo do
        // cabeçalho — que no arquivo tem posição de verdade — caía no centro.
        var contentWidthEmus = Math.Max(
            (widthTwips - (margin?.Left?.Value ?? 1440) - (margin?.Right?.Value ?? 1440)) * EmusPerTwip,
            1);

        return new PageSetupDto(
            Size: NearestSize(widthTwips, heightTwips, landscape),
            Orientation: landscape ? "landscape" : "portrait",
            Margins: new MarginsDto(
                Top: Millimeters(margin?.Top?.Value, 1440),
                Right: Millimeters((int?)margin?.Right?.Value, 1440),
                Bottom: Millimeters(margin?.Bottom?.Value, 1440),
                Left: Millimeters((int?)margin?.Left?.Value, 1440)),
            Header: NullIfEmpty(HeaderReader.Read(section, part, inventory, HeaderFooterValues.Default, contentWidthEmus)),
            Footer: NullIfEmpty(HeaderReader.ReadFooter(section, part, inventory, HeaderFooterValues.Default, contentWidthEmus)),
            // Os dois interruptores decidem se as referências valem. O Word
            // guarda o `first` mesmo com `w:titlePg` desligado — usá-lo sem
            // conferir poria a capa em todas as páginas.
            FirstHeader: HasTitlePage(section)
                ? NullIfEmpty(HeaderReader.Read(section, part, inventory, HeaderFooterValues.First, contentWidthEmus))
                : null,
            FirstFooter: HasTitlePage(section)
                ? NullIfEmpty(HeaderReader.ReadFooter(section, part, inventory, HeaderFooterValues.First, contentWidthEmus))
                : null,
            EvenHeader: UsesEvenAndOdd(part)
                ? NullIfEmpty(HeaderReader.Read(section, part, inventory, HeaderFooterValues.Even, contentWidthEmus))
                : null,
            EvenFooter: UsesEvenAndOdd(part)
                ? NullIfEmpty(HeaderReader.ReadFooter(section, part, inventory, HeaderFooterValues.Even, contentWidthEmus))
                : null);
    }

    /// <summary>`w:titlePg`: a primeira página tem cabeçalho próprio.</summary>
    /// <remarks>
    /// Elemento presente sem `w:val` significa ligado — é a convenção dos
    /// interruptores do OOXML, e lê-lo como desligado é o erro clássico.
    /// </remarks>
    private static bool HasTitlePage(SectionProperties section)
    {
        var flag = section.GetFirstChild<TitlePage>();
        return flag is not null && (flag.Val?.Value ?? true);
    }

    /// <summary>
    /// `w:evenAndOddHeaders`: páginas pares têm cabeçalho próprio.
    /// </summary>
    /// <remarks>
    /// Mora em `settings.xml`, e não na seção: no Word é escolha do documento
    /// inteiro, não de um trecho dele.
    /// </remarks>
    private static bool UsesEvenAndOdd(MainDocumentPart part)
    {
        var flag = part.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>();
        return flag is not null && (flag.Val?.Value ?? true);
    }

    private static bool AllShareGeometry(List<SectionProperties> sections)
    {
        static (int W, int H, int T, int R, int B, int L) Geometry(SectionProperties section)
        {
            var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
            var margin = section.GetFirstChild<PageMargin>();
            return (
                (int?)size?.Width?.Value ?? 0,
                (int?)size?.Height?.Value ?? 0,
                margin?.Top?.Value ?? 0,
                (int?)margin?.Right?.Value ?? 0,
                margin?.Bottom?.Value ?? 0,
                (int?)margin?.Left?.Value ?? 0);
        }

        var first = Geometry(sections[0]);
        return sections.Skip(1).Select(Geometry).All(other =>
            Math.Abs(other.W - first.W) <= GeometryTolerance &&
            Math.Abs(other.H - first.H) <= GeometryTolerance &&
            Math.Abs(other.T - first.T) <= GeometryTolerance &&
            Math.Abs(other.R - first.R) <= GeometryTolerance &&
            Math.Abs(other.B - first.B) <= GeometryTolerance &&
            Math.Abs(other.L - first.L) <= GeometryTolerance);
    }

    /// <summary>Faixa vazia vira ausência: o modelo distingue "não tem" de "tem e está vazia".</summary>
    private static BandDto? NullIfEmpty(BandDto band) => band.IsEmpty ? null : band;

    private static double Millimeters(int? twips, int fallback) =>
        Math.Round((twips ?? fallback) / TwipsPerMillimeter, 1);

    /// <summary>
    /// O modelo só conhece A4 e Carta. Um tamanho fora disso vira o mais
    /// próximo — preferível a recusar o documento por causa do papel.
    /// </summary>
    private static string NearestSize(double widthTwips, double heightTwips, bool landscape)
    {
        var shortSide = landscape ? heightTwips : widthTwips;
        // A4 tem 11906 twips de largura; Carta tem 12240.
        return Math.Abs(shortSide - 11906) <= Math.Abs(shortSide - 12240) ? "A4" : "Letter";
    }

    private static PageSetupDto Default() => new(
        "A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null);
}

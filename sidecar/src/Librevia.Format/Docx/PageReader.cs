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
    [property: JsonPropertyName("evenFooterBand")] BandDto? EvenFooter = null,
    /// <summary>
    /// Distância da faixa à borda do papel (`w:pgMar/@header` e `@footer`).
    /// </summary>
    /// <remarks>
    /// É a origem vertical das âncoras de dentro do cabeçalho: elas se dizem
    /// relativas ao "parágrafo", e o parágrafo do cabeçalho começa justamente
    /// aqui. Sem esta medida, um objeto ancorado na faixa não tem de onde contar.
    /// </remarks>
    [property: JsonPropertyName("headerDistanceMm")] double HeaderDistanceMm = 12.5,
    [property: JsonPropertyName("footerDistanceMm")] double FooterDistanceMm = 12.5);

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

    /// <summary>
    /// Meio milímetro de folga ao reconhecer o papel.
    /// </summary>
    /// <remarks>
    /// As medidas do papel variam no último twip entre quem grava o arquivo, e A4
    /// gravado pelo LibreOffice e pelo Word não bate no dígito.
    /// </remarks>
    private const int PaperTolerance = 30;

    /// <summary>
    /// Os papéis que o modelo do editor nomeia, com as medidas em twips.
    /// </summary>
    /// <remarks>
    /// Uma tabela só. As medidas estavam em três lugares — aqui, em milímetros na
    /// conta da coluna de texto e outra vez em twips na hora de gravar o
    /// `w:pgSz` —, e três cópias do mesmo número é como um A4 passa a ter duas
    /// larguras.
    /// </remarks>
    private static readonly (string Name, uint Short, uint Long)[] Papers =
    [
        ("A4", 11906U, 16838U),
        ("Letter", 12240U, 15840U),
    ];

    /// <summary>
    /// As medidas em twips do papel que o modelo nomeia; A4 para um nome que a
    /// tabela não tem.
    /// </summary>
    public static (uint Short, uint Long) TwipsOfPaper(string name)
    {
        foreach (var paper in Papers)
        {
            if (string.Equals(paper.Name, name, StringComparison.Ordinal)) return (paper.Short, paper.Long);
        }

        return (Papers[0].Short, Papers[0].Long);
    }

    /// <summary>As mesmas medidas em milímetros, que é a unidade do modelo.</summary>
    public static (double Short, double Long) MillimetersOfPaper(string name)
    {
        var (shortSide, longSide) = TwipsOfPaper(name);
        return (shortSide / TwipsPerMillimeter, longSide / TwipsPerMillimeter);
    }

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
        // Sem `w:pgSz` valem as medidas do primeiro papel da tabela: é o que o
        // painel mostraria de qualquer modo.
        var widthTwips = (double?)size?.Width?.Value ?? Papers[0].Short;
        var heightTwips = (double?)size?.Height?.Value ?? Papers[0].Long;

        // O painel de página só oferece A4 e Carta, e um papel fora dos dois é
        // mostrado como o mais próximo. O arquivo **mantém** a medida dele — a
        // gravação não regrava `w:pgSz` sem necessidade —, mas o usuário veria A4
        // escrito na tela sem saber de onde veio. Invisibilidade, não perda.
        if (!IsKnownPaper(widthTwips, heightTwips))
        {
            inventory.NoteInvisible(
                "o tamanho do papel deste documento não é A4 nem Carta (ele é preservado no arquivo)");
        }

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
                : null,
            HeaderDistanceMm: Millimeters((int?)margin?.Header?.Value, 708),
            FooterDistanceMm: Millimeters((int?)margin?.Footer?.Value, 708));
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
    /// O papel é um dos que o modelo nomeia?
    /// </summary>
    private static bool IsKnownPaper(double widthTwips, double heightTwips)
    {
        var shortSide = Math.Min(widthTwips, heightTwips);
        var longSide = Math.Max(widthTwips, heightTwips);

        return Papers.Any(paper =>
            Math.Abs(shortSide - paper.Short) <= PaperTolerance
            && Math.Abs(longSide - paper.Long) <= PaperTolerance);
    }

    /// <summary>
    /// O modelo só conhece A4 e Carta. Um tamanho fora disso vira o mais
    /// próximo — preferível a recusar o documento por causa do papel.
    /// </summary>
    /// <remarks>
    /// Compara o lado que **atravessa** a folha, que é o que separa os dois: eles
    /// têm alturas bem diferentes e larguras parecidas. Empate fica com o
    /// primeiro da tabela.
    /// </remarks>
    private static string NearestSize(double widthTwips, double heightTwips, bool landscape)
    {
        var across = landscape ? heightTwips : widthTwips;
        return Papers.MinBy(paper => Math.Abs(across - paper.Short)).Name;
    }

    /// <summary>
    /// A configuração que o modelo traz é a mesma que esta seção já declara?
    /// </summary>
    /// <remarks>
    /// Serve para **não** regravar o `w:sectPr` de quem não mexeu na página. A
    /// gravação o regravava sempre, e com isso todo papel fora de A4 e Carta era
    /// arredondado para um dos dois na primeira vez que se salvasse o arquivo: um
    /// documento em A5 virava A4 por ter recebido uma correção de vírgula.
    ///
    /// A comparação é feita nas unidades do modelo — nome do papel, orientação e
    /// margens em décimos de milímetro —, e não em twips: é o modelo que dá a
    /// resolução com que o usuário pode ter mudado algo.
    /// </remarks>
    public static bool Matches(SectionProperties section, PageSetupDto page)
    {
        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        var margin = section.GetFirstChild<PageMargin>();

        var landscape = size?.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        // Sem `w:pgSz` valem as medidas do primeiro papel da tabela: é o que o
        // painel mostraria de qualquer modo.
        var widthTwips = (double?)size?.Width?.Value ?? Papers[0].Short;
        var heightTwips = (double?)size?.Height?.Value ?? Papers[0].Long;

        return landscape == string.Equals(page.Orientation, "landscape", StringComparison.Ordinal)
               && string.Equals(NearestSize(widthTwips, heightTwips, landscape), page.Size, StringComparison.Ordinal)
               && Millimeters(margin?.Top?.Value, 1440) == page.Margins.Top
               && Millimeters((int?)margin?.Right?.Value, 1440) == page.Margins.Right
               && Millimeters(margin?.Bottom?.Value, 1440) == page.Margins.Bottom
               && Millimeters((int?)margin?.Left?.Value, 1440) == page.Margins.Left;
    }

    /// <summary>
    /// O nome que o modelo daria a um papel destas medidas.
    /// </summary>
    /// <remarks>
    /// Público para a gravação: é assim que ela sabe se as medidas do arquivo
    /// ainda **são** o papel que o modelo diz, e pode então preservá-las em vez
    /// de trocá-las pelas medidas canônicas. É o que mantém um A5 sendo A5
    /// depois de a pessoa virar a folha para paisagem.
    /// </remarks>
    public static string NameOfPaper(uint? widthTwips, uint? heightTwips, bool landscape) =>
        NearestSize(widthTwips ?? Papers[0].Short, heightTwips ?? Papers[0].Long, landscape);

    private static PageSetupDto Default() => new(
        "A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null);
}

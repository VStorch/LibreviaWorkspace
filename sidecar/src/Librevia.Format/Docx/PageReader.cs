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
    [property: JsonPropertyName("footerDistanceMm")] double FooterDistanceMm = 12.5,
    // O cabeçalho e o rodapé de texto simples do documento novo, com `{n}` e
    // `{total}` no lugar dos números. Quem os grava é PlainBandWriter.
    [property: JsonPropertyName("header")] string? HeaderText = null,
    [property: JsonPropertyName("footer")] string? FooterText = null,
    // Numeração de página (`w:pgNumType`) e os dois interruptores das faixas.
    // Anuláveis: um `.sdoc` antigo não os traz, e ausência quer dizer "não mexa
    // no que o arquivo já diz" — e não "desligue".
    [property: JsonPropertyName("pageNumberFormat")] string? PageNumberFormat = null,
    // Ausente e nulo são coisas diferentes aqui: ausente (rascunho de antes) é
    // "não mexa no `w:start`"; nulo é "sem início", e apaga o que houver. Por isso
    // `JsonElement`, que distingue os dois — ver PageNumberStartOf.
    [property: JsonPropertyName("pageNumberStart")] System.Text.Json.JsonElement PageNumberStart = default,
    [property: JsonPropertyName("titlePage")] bool? TitlePage = null,
    [property: JsonPropertyName("evenAndOddHeaders")] bool? EvenAndOddHeaders = null,
    // Seções (M9). `Id` só existe nas seções anteriores à última: é o valor do
    // atributo `sectionBreak` do parágrafo que carrega o `w:sectPr` delas. A
    // última é o `w:sectPr` do corpo, e não tem parágrafo nem id.
    [property: JsonPropertyName("id")] string? Id = null,
    // Como a seção começa (`w:sectPr/w:type`): nextPage, continuous, evenPage,
    // oddPage ou nextColumn. Ausente (rascunho de antes) é "não mexa".
    [property: JsonPropertyName("start")] string? Start = null);

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

    /// <summary>
    /// As seções do documento: a última (o `w:sectPr` do corpo) e as anteriores.
    /// </summary>
    /// <remarks>
    /// A última continua sendo "a página" do modelo — é ela que o documento de uma
    /// seção só tem, e é por ela que o `.sdoc` de antes das seções continua
    /// abrindo igual. As anteriores vêm em ordem, cada uma com o id que o leitor
    /// do corpo põe no parágrafo que a encerra (ver <see cref="SectionIds"/>).
    ///
    /// Cada seção leva só as faixas que ela **declara**. A que não declara um tipo
    /// herda o da seção anterior — é o "Vincular ao anterior" do Word —, e quem
    /// resolve a herança é quem desenha: repetir a faixa em cada seção faria duas
    /// cópias de uma parte só, e editar uma não mudaria a outra.
    /// </remarks>
    public static (PageSetupDto Page, List<PageSetupDto>? Sections) ReadAll(
        Body body,
        MainDocumentPart part,
        Inventory inventory)
    {
        var all = body.Descendants<SectionProperties>().ToList();
        if (all.Count == 0) return (Default(), null);

        // A última é a do corpo; as de parágrafo vêm antes, na ordem do arquivo.
        // O documento que termina num `w:sectPr` de parágrafo sem o do corpo é
        // inválido, mas acontece: a última de parágrafo faz as vezes da final.
        var ids = SectionIds(body);
        var earlier = new List<PageSetupDto>();
        for (var index = 0; index < all.Count - 1; index++)
        {
            var section = all[index];
            earlier.Add(ReadOne(section, part, inventory, index) with
            {
                Id = ids.TryGetValue(section, out var id) ? id : $"s{index + 1}",
            });
        }

        var page = ReadOne(all[^1], part, inventory, all.Count - 1);
        return (page, earlier.Count == 0 ? null : earlier);
    }

    /// <summary>
    /// O id de cada `w:sectPr` de parágrafo: `s1`, `s2`… na ordem do corpo.
    /// </summary>
    /// <remarks>
    /// Posicional e determinístico, como o `oid` dos blocos: a leitura de
    /// referência da gravação reproduz os mesmos ids, e é isso que deixa o
    /// parágrafo da marca ser reconhecido e voltar byte a byte.
    /// </remarks>
    public static Dictionary<SectionProperties, string> SectionIds(Body body)
    {
        var ids = new Dictionary<SectionProperties, string>(ReferenceEqualityComparer.Instance);
        var all = body.Descendants<SectionProperties>().ToList();
        for (var index = 0; index < all.Count - 1; index++) ids[all[index]] = $"s{index + 1}";
        return ids;
    }

    private static PageSetupDto ReadOne(SectionProperties section, MainDocumentPart part, Inventory inventory, int index)
    {
        // Formato de número que a tela não desenha: a folha mostra decimal, e o
        // arquivo continua pedindo o dele — é diferença de aparência, e se avisa.
        if (section.GetFirstChild<PageNumberType>()?.Format?.InnerText is { } pageFormat &&
            !PageNumberFormats.Contains(pageFormat))
        {
            inventory.NoteInvisible(
                $"formato de número de página \"{pageFormat}\" (mostrado em algarismos; o arquivo o mantém)");
        }
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

        // As faixas de capa e de página par vêm sempre que o arquivo as tem;
        // quem decide se valem são os interruptores, que vão junto. O Word
        // guarda o `first` mesmo com `w:titlePg` desligado — e é justamente
        // ele que volta a aparecer quando a pessoa liga o interruptor na
        // configuração de página. Lido só com o interruptor ligado, ligar
        // mostrava a capa em branco na tela e a do arquivo no Word.
        BandDto? Band(bool header, HeaderFooterValues type)
        {
            // Da segunda seção em diante, faixa ausente é faixa herdada — e a
            // declarada vazia é folha limpa, não herança: fica vazia, e não nula.
            if (index > 0 && !Declares(section, header, type)) return null;
            var band = header
                ? HeaderReader.Read(section, part, inventory, type, contentWidthEmus)
                : HeaderReader.ReadFooter(section, part, inventory, type, contentWidthEmus);
            return index == 0 ? NullIfEmpty(band) : band;
        }

        return new PageSetupDto(
            Size: NearestSize(widthTwips, heightTwips, landscape),
            Orientation: landscape ? "landscape" : "portrait",
            Margins: new MarginsDto(
                Top: Millimeters(margin?.Top?.Value, 1440),
                Right: Millimeters((int?)margin?.Right?.Value, 1440),
                Bottom: Millimeters(margin?.Bottom?.Value, 1440),
                Left: Millimeters((int?)margin?.Left?.Value, 1440)),
            Header: Band(true, HeaderFooterValues.Default),
            Footer: Band(false, HeaderFooterValues.Default),
            FirstHeader: Band(true, HeaderFooterValues.First),
            FirstFooter: Band(false, HeaderFooterValues.First),
            EvenHeader: Band(true, HeaderFooterValues.Even),
            EvenFooter: Band(false, HeaderFooterValues.Even),
            HeaderDistanceMm: Millimeters((int?)margin?.Header?.Value, 708),
            FooterDistanceMm: Millimeters((int?)margin?.Footer?.Value, 708),
            PageNumberFormat: PageNumberFormatOf(section),
            PageNumberStart: StartElement(section.GetFirstChild<PageNumberType>()?.Start?.Value),
            TitlePage: HasTitlePage(section),
            EvenAndOddHeaders: UsesEvenAndOdd(part),
            Start: StartOf(section));
    }

    /// <summary>A seção declara uma faixa deste tipo, ainda que vazia?</summary>
    private static bool Declares(SectionProperties section, bool header, HeaderFooterValues type)
    {
        IEnumerable<HeaderFooterReferenceType> references = header
            ? section.Elements<HeaderReference>()
            : section.Elements<FooterReference>();
        return references.Any(reference =>
            (reference.Type?.Value ?? HeaderFooterValues.Default) == type &&
            !string.IsNullOrEmpty(reference.Id?.Value));
    }

    /// <summary>Os começos de seção que o modelo nomeia — os de `w:type/@w:val`.</summary>
    public static readonly string[] SectionStarts = ["nextPage", "continuous", "evenPage", "oddPage", "nextColumn"];

    /// <summary>
    /// `w:type/@w:val`, ou "nextPage" — o padrão da especificação quando o
    /// elemento falta.
    /// </summary>
    public static string StartOf(SectionProperties section)
    {
        var name = section.GetFirstChild<SectionType>()?.Val?.InnerText;
        return name is not null && SectionStarts.Contains(name) ? name : "nextPage";
    }

    /// <summary>O início como o modelo o leva: número ou nulo, sempre presente.</summary>
    public static System.Text.Json.JsonElement StartElement(int? start) =>
        System.Text.Json.JsonSerializer.SerializeToElement(start);

    /// <summary>
    /// O início que o modelo pede; `false` quando ele não diz nada (campo ausente).
    /// </summary>
    public static bool TryStartOf(PageSetupDto page, out int? start)
    {
        start = null;
        switch (page.PageNumberStart.ValueKind)
        {
            case System.Text.Json.JsonValueKind.Number when page.PageNumberStart.TryGetInt32(out var value):
                start = value;
                return true;
            case System.Text.Json.JsonValueKind.Null:
                return true;
            default:
                return false;
        }
    }

    /// <summary>Os formatos de número de página que o editor desenha.</summary>
    public static readonly string[] PageNumberFormats =
        ["decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter"];

    /// <summary>
    /// `w:pgNumType/@w:fmt`, ou decimal.
    /// </summary>
    /// <remarks>
    /// Formato que o editor não desenha (`numberInDash`, os de outros alfabetos)
    /// vira decimal na tela, mas volta intacto ao arquivo: a gravação só toca o
    /// atributo quando o modelo diz algo diferente do que se leu.
    /// </remarks>
    public static string PageNumberFormatOf(SectionProperties section)
    {
        var name = section.GetFirstChild<PageNumberType>()?.Format?.InnerText;
        return name is not null && PageNumberFormats.Contains(name) ? name : "decimal";
    }

    /// <inheritdoc cref="HasTitlePage(SectionProperties)"/>
    public static bool TitlePageOf(SectionProperties section) => HasTitlePage(section);

    /// <inheritdoc cref="UsesEvenAndOdd(MainDocumentPart)"/>
    public static bool EvenAndOddOf(MainDocumentPart part) => UsesEvenAndOdd(part);

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
        "A4", "portrait", new MarginsDto(25, 25, 25, 25), null, null,
        PageNumberFormat: "decimal", TitlePage: false, EvenAndOddHeaders: false);
}

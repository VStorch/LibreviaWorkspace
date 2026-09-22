using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O `word/styles.xml` do documento novo.
/// </summary>
/// <remarks>
/// Não é um gosto: é a **tela**. O editor desenha o documento novo com o CSS de
/// `src/services/document/content-styles.ts` e com o padrão do navegador para o
/// que o CSS não diz, e o arquivo tem de reproduzir exatamente isso — se tela e
/// arquivo discordassem, salvar em DOCX mudaria a paginação do que a pessoa
/// acabou de escrever. Os números abaixo foram medidos no editor (estilo
/// computado de `p` e de `h1` a `h6` num documento novo), e não copiados do Word.
///
/// Quando os estilos virarem dado, estes números saem daqui e passam a vir do
/// documento; até lá, mudar a aparência é mudar os dois lugares juntos.
///
/// Duas aproximações, ambas abaixo do que o olho vê: o tamanho do `h5` e do `h6`
/// (9,96 pt e 8,04 pt no navegador) vai para o meio-ponto mais próximo, que é a
/// precisão do `w:sz`; e o `#111111` do texto fica `auto`, que é preto — gravado
/// como cor, ele voltaria como cor explícita em cada trecho reaberto.
/// </remarks>
internal static class TemplateStyles
{
    /// <summary>A fonte do documento novo — a primeira da pilha do CSS.</summary>
    public const string BodyFont = "Times New Roman";

    /// <summary>A fonte do cabeçalho e do rodapé de texto simples (`page-setup.ts`).</summary>
    public const string BandFont = "Calibri";

    /// <summary>
    /// Entrelinha 1,5 do CSS, em 240-avos da altura natural da fonte.
    /// </summary>
    /// <remarks>
    /// A mesma conta de <c>ParagraphFormat.ApplyLineHeight</c>: o múltiplo do
    /// OOXML é medido sobre a altura da fonte, e não sobre o tamanho dela.
    /// </remarks>
    private static readonly int BodyLine = (int)Math.Round(1.5 / (LineMetrics.Of(BodyFont) ?? 1.1499) * 240);

    /// <summary>
    /// Os títulos: tamanho em meios-pontos, espaço antes e depois em twips e se
    /// fica preso ao parágrafo seguinte.
    /// </summary>
    /// <remarks>
    /// O antes é o `margin-top: 1em` do CSS (0,6em no `h5` e no `h6`, que não têm
    /// regra própria); o depois é a margem de baixo do navegador — 0,67em no
    /// `h1`, 0,83em no `h2`, 1em no `h3`, 1,33em no `h4`, 1,67em e 2,33em nos dois
    /// últimos. O "manter com o próximo" é o `break-after: avoid` da impressão,
    /// que só os quatro primeiros têm.
    /// </remarks>
    private static readonly (int HalfPoints, int Before, int After, bool KeepNext)[] Headings =
    [
        (44, 440, 295, true),
        (34, 340, 282, true),
        (28, 280, 280, true),
        (24, 240, 319, true),
        (20, 120, 333, false),
        (16, 96, 375, false),
    ];

    public static Styles Create()
    {
        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    // Os quatro atributos, e por nome: não há tema no pacote, e
                    // uma letra acentuada ou um caractere asiático cairia na fonte
                    // que o programa quisesse.
                    new RunFonts { Ascii = BodyFont, HighAnsi = BodyFont, EastAsia = BodyFont, ComplexScript = BodyFont },
                    new FontSize { Val = "24" },
                    new FontSizeComplexScript { Val = "24" })),
                new ParagraphPropertiesDefault()),
            Normal(),
            DefaultParagraphFont(),
            TableNormal(),
            NoList());

        for (var level = 1; level <= Headings.Length; level++) styles.AppendChild(Heading(level));

        styles.AppendChild(TableGrid());
        styles.AppendChild(ListParagraph());
        styles.AppendChild(Hyperlink());
        return styles;
    }

    /// <summary>A definição de um título, pronta para ir a outro pacote.</summary>
    /// <remarks>
    /// Pública para <see cref="HeadingStyles"/>: o título criado num DOCX que não
    /// define o estilo leva **esta** definição, e não outra — duas cópias da
    /// mesma aparência divergiriam na primeira correção.
    /// </remarks>
    public static Style Heading(int level)
    {
        var (halfPoints, before, after, keepNext) = Headings[level - 1];

        var paragraph = new StyleParagraphProperties();
        if (keepNext) paragraph.AppendChild(new KeepNext());
        paragraph.AppendChild(new SpacingBetweenLines { Before = Invariant(before), After = Invariant(after) });
        paragraph.AppendChild(new OutlineLevel { Val = level - 1 });

        return new Style(
            new StyleName { Val = $"heading {level}" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new UIPriority { Val = 9 },
            new PrimaryStyle(),
            paragraph,
            new StyleRunProperties(
                new Bold(),
                new BoldComplexScript(),
                new FontSize { Val = Invariant(halfPoints) },
                new FontSizeComplexScript { Val = Invariant(halfPoints) }))
        {
            Type = StyleValues.Paragraph,
            StyleId = $"Heading{level}",
        };
    }

    public static int HeadingLevels => Headings.Length;

    private static Style Normal() => new(
        new StyleName { Val = "Normal" },
        new PrimaryStyle(),
        new StyleParagraphProperties(new SpacingBetweenLines
        {
            // 0,6em antes (`.page__content > * + *`) e 1em depois (o padrão do
            // navegador para `p`), sobre 12 pt.
            Before = "144",
            After = "240",
            Line = Invariant(BodyLine),
            LineRule = LineSpacingRuleValues.Auto,
        }))
    {
        Type = StyleValues.Paragraph,
        StyleId = "Normal",
        Default = true,
    };

    private static Style DefaultParagraphFont() => new(
        new StyleName { Val = "Default Paragraph Font" },
        new UIPriority { Val = 1 },
        new SemiHidden(),
        new UnhideWhenUsed())
    {
        Type = StyleValues.Character,
        StyleId = "DefaultParagraphFont",
        Default = true,
    };

    /// <remarks>
    /// A margem da célula é o `padding: 4px 8px` do CSS: 3 pt em cima e embaixo,
    /// 6 pt dos lados.
    /// </remarks>
    private static Style TableNormal() => new(
        new StyleName { Val = "Normal Table" },
        new UIPriority { Val = 99 },
        new SemiHidden(),
        new UnhideWhenUsed(),
        new StyleTableProperties(
            new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa },
            new TableCellMarginDefault(
                new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 120, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 120, Type = TableWidthValues.Dxa })))
    {
        Type = StyleValues.Table,
        StyleId = "TableNormal",
        Default = true,
    };

    private static Style NoList() => new(
        new StyleName { Val = "No List" },
        new UIPriority { Val = 99 },
        new SemiHidden(),
        new UnhideWhenUsed())
    {
        Type = StyleValues.Numbering,
        StyleId = "NoList",
        Default = true,
    };

    private static Style TableGrid()
    {
        var val = BorderValues.Single;
        const uint size = 4;

        return new Style(
            new StyleName { Val = "Table Grid" },
            new BasedOn { Val = "TableNormal" },
            new UIPriority { Val = 39 },
            new StyleParagraphProperties(new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new StyleTableProperties(new TableBorders(
                new TopBorder { Val = val, Size = size, Space = 0U, Color = "auto" },
                new LeftBorder { Val = val, Size = size, Space = 0U, Color = "auto" },
                new BottomBorder { Val = val, Size = size, Space = 0U, Color = "auto" },
                new RightBorder { Val = val, Size = size, Space = 0U, Color = "auto" },
                new InsideHorizontalBorder { Val = val, Size = size, Space = 0U, Color = "auto" },
                new InsideVerticalBorder { Val = val, Size = size, Space = 0U, Color = "auto" })))
        {
            Type = StyleValues.Table,
            StyleId = "TableGrid",
        };
    }

    private static Style ListParagraph() => new(
        new StyleName { Val = "List Paragraph" },
        new BasedOn { Val = "Normal" },
        new UIPriority { Val = 34 },
        new PrimaryStyle(),
        new StyleParagraphProperties(new Indentation { Left = "720" }, new ContextualSpacing()))
    {
        Type = StyleValues.Paragraph,
        StyleId = "ListParagraph",
    };

    private static Style Hyperlink() => new(
        new StyleName { Val = "Hyperlink" },
        new BasedOn { Val = "DefaultParagraphFont" },
        new UIPriority { Val = 99 },
        new UnhideWhenUsed(),
        new StyleRunProperties(new Color { Val = "0563C1" }, new Underline { Val = UnderlineValues.Single }))
    {
        Type = StyleValues.Character,
        StyleId = "Hyperlink",
    };

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);
}

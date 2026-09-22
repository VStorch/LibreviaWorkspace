using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O `word/styles.xml` do documento novo, montado a partir de <see cref="BuiltinStyles"/>.
/// </summary>
/// <remarks>
/// Aqui só a **montagem**: os números moram na tabela de dados, nas unidades do
/// editor, e é ela que o teste de contrato compara com o `BUILTIN_STYLES` do lado
/// TS. Enquanto os dois viviam juntos, mudar a aparência pedia mudar twips e
/// meios-pontos em dois idiomas, e a chance de um lado ficar atrás era certa.
///
/// A tabela não cobre os estilos de tabela e de numeração (`TableNormal`,
/// `NoList`, `TableGrid`): o modelo do documento não representa borda nem margem
/// de célula, então eles continuam montados à mão logo abaixo.
/// </remarks>
internal static class TemplateStyles
{
    /// <inheritdoc cref="BuiltinStyles.BodyFont"/>
    public const string BodyFont = BuiltinStyles.BodyFont;

    /// <inheritdoc cref="BuiltinStyles.BandFont"/>
    public const string BandFont = BuiltinStyles.BandFont;

    public static Styles Create()
    {
        var bodySize = HalfPoints(BuiltinStyles.BodySizePt);

        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    // Os quatro atributos, e por nome: não há tema no pacote, e
                    // uma letra acentuada ou um caractere asiático cairia na fonte
                    // que o programa quisesse.
                    new RunFonts { Ascii = BodyFont, HighAnsi = BodyFont, EastAsia = BodyFont, ComplexScript = BodyFont },
                    new FontSize { Val = bodySize },
                    new FontSizeComplexScript { Val = bodySize })),
                new ParagraphPropertiesDefault()));

        // Os de parágrafo e de caractere saem da tabela, na ordem dela; os de
        // tabela e de numeração entram no meio, onde estavam: `TableNormal` e
        // `NoList` são os padrões do pacote e o Word os espera antes do resto.
        foreach (var style in BuiltinStyles.All)
        {
            if (style.Id == "Heading1")
            {
                styles.AppendChild(TableNormal());
                styles.AppendChild(NoList());
            }

            styles.AppendChild(Of(style));

            if (style.Id == "Heading6") styles.AppendChild(TableGrid());
        }

        return styles;
    }

    /// <summary>A definição de um título, pronta para ir a outro pacote.</summary>
    /// <remarks>
    /// Pública para <see cref="HeadingStyles"/>: o título criado num DOCX que não
    /// define o estilo leva **esta** definição, e não outra — duas cópias da
    /// mesma aparência divergiriam na primeira correção.
    /// </remarks>
    public static Style Heading(int level) => Of(BuiltinStyles.Heading(level));

    public static int HeadingLevels => BuiltinStyles.HeadingLevels;

    /// <summary>
    /// Um estilo da tabela de dados, em OOXML.
    /// </summary>
    /// <remarks>
    /// A ordem dos filhos não é gosto: o esquema do OOXML a fixa (nome, herança,
    /// seguinte, ligado, escondido, prioridade, qFormat, e só então `w:pPr` e
    /// `w:rPr`), e fora dela o Word recusa o arquivo. Quem confere é o
    /// `OpenXmlValidator`, em todo DOCX que os testes gravam.
    /// </remarks>
    private static Style Of(BuiltinStyle style)
    {
        var result = new Style
        {
            Type = style.Character ? StyleValues.Character : StyleValues.Paragraph,
            StyleId = style.Id,
        };

        result.AppendChild(new StyleName { Val = style.Name });
        if (style.BasedOn is not null) result.AppendChild(new BasedOn { Val = style.BasedOn });
        if (style.Next is not null) result.AppendChild(new NextParagraphStyle { Val = style.Next });
        if (style.Link is not null) result.AppendChild(new LinkedStyle { Val = style.Link });
        if (style.Hidden) result.AppendChild(new StyleHidden());
        if (style.UiPriority is { } priority) result.AppendChild(new UIPriority { Val = priority });
        if (style.SemiHidden) result.AppendChild(new SemiHidden());
        if (style.UnhideWhenUsed) result.AppendChild(new UnhideWhenUsed());
        if (style.QFormat) result.AppendChild(new PrimaryStyle());
        if (style.Default) result.Default = true;

        if (ParagraphOf(style) is { } paragraph) result.AppendChild(paragraph);
        if (RunOf(style) is { } run) result.AppendChild(run);

        return result;
    }

    private static StyleParagraphProperties? ParagraphOf(BuiltinStyle style)
    {
        var properties = new StyleParagraphProperties();

        if (style.KeepNext) properties.AppendChild(new KeepNext());

        if (style.BeforePt is not null || style.AfterPt is not null || style.LineFactor is not null)
        {
            var spacing = new SpacingBetweenLines();
            if (style.BeforePt is { } before) spacing.Before = Twips(before);
            if (style.AfterPt is { } after) spacing.After = Twips(after);
            if (style.LineFactor is { } factor)
            {
                // O múltiplo do OOXML vem em 240-avos da altura natural da linha
                // — a mesma conta de `ParagraphFormat.ApplyLineHeight`.
                spacing.Line = Invariant((int)Math.Round(factor * 240));
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }

            properties.AppendChild(spacing);
        }

        if (style.IndentMm is { } indent)
        {
            properties.AppendChild(new Indentation { Left = Invariant(Attr.MmToTwips(indent)) });
        }

        if (style.ContextualSpacing) properties.AppendChild(new ContextualSpacing());
        if (style.OutlineLevel is { } level) properties.AppendChild(new OutlineLevel { Val = level });

        return properties.HasChildren ? properties : null;
    }

    private static StyleRunProperties? RunOf(BuiltinStyle style)
    {
        var properties = new StyleRunProperties();

        // O negrito e o tamanho vão também na variante de escrita complexa: sem
        // ela, um trecho em árabe ou hebraico dentro do título sai sem negrito e
        // no tamanho do corpo.
        if (style.Bold)
        {
            properties.AppendChild(new Bold());
            properties.AppendChild(new BoldComplexScript());
        }

        if (style.Color is { } color)
        {
            properties.AppendChild(new Color { Val = color.TrimStart('#').ToUpperInvariant() });
        }

        if (style.SizePt is { } size)
        {
            properties.AppendChild(new FontSize { Val = HalfPoints(size) });
            properties.AppendChild(new FontSizeComplexScript { Val = HalfPoints(size) });
        }

        if (style.Underline) properties.AppendChild(new Underline { Val = UnderlineValues.Single });

        return properties.HasChildren ? properties : null;
    }

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

    /// <summary>Pontos → meios-pontos, que é a unidade do `w:sz`.</summary>
    private static string HalfPoints(double points) => Invariant((int)Math.Round(points * 2));

    /// <summary>Pontos → twips, que é a unidade do `w:spacing`.</summary>
    private static string Twips(double points) => Invariant((int)Math.Round(points * 20));

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);
}

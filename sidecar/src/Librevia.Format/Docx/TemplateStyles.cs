using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// The new document's <c>word/styles.xml</c>. The numbers live in the data table (<see
/// cref="BuiltinStyles"/>); table and numbering styles are built here.
/// </summary>
internal static class TemplateStyles
{
    /// <inheritdoc cref="BuiltinStyles.BodyFont"/>
    public const string BodyFont = BuiltinStyles.BodyFont;

    /// <inheritdoc cref="BuiltinStyles.BandFont"/>
    public const string BandFont = BuiltinStyles.BandFont;

    /// <summary>
    /// The template's styles, or the <see cref="BuiltinStyles"/> table: a document born in the
    /// editor is Calibri, and an old <c>.sdoc</c>, Times. A style the user created belongs to <see
    /// cref="StyleWriter"/>.
    /// </summary>
    public static Styles Create(StyleSheetDto? sheet = null)
    {
        var font = FirstFontOf(sheet?.Defaults.Character.FontFamily) ?? BodyFont;
        var bodySize = HalfPoints(PointsOf(sheet?.Defaults.Character.FontSize) ?? BuiltinStyles.BodySizePt);

        var paragraphDefault = new ParagraphPropertiesDefault();
        if (sheet?.Defaults.Paragraph is { } defaults &&
            SpacingOf(defaults.SpaceBefore, defaults.SpaceAfter, MultipleOf(defaults.LineSpacing)) is { } spacing)
        {
            paragraphDefault.AppendChild(new ParagraphPropertiesBaseStyle(spacing));
        }

        var styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    // All four attributes: without a theme, accents and Asian scripts would fall
                    // into another font.
                    new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font },
                    new FontSize { Val = bodySize },
                    new FontSizeComplexScript { Val = bodySize })),
                paragraphDefault));

        // `TableNormal` and `NoList` are the package defaults, and Word expects them before the
        // rest.
        var table = sheet is null ? BuiltinStyles.All : [.. sheet.Styles.Values.Select(style => FromSheet(style, sheet.Defaults))];
        var packageDefaults = false;
        var grid = false;
        foreach (var style in table)
        {
            if (!packageDefaults && style.Id == "Heading1")
            {
                styles.AppendChild(TableNormal());
                styles.AppendChild(NoList());
                packageDefaults = true;
            }

            styles.AppendChild(Of(style));

            if (style.Id == "Heading6")
            {
                styles.AppendChild(TableGrid());
                grid = true;
            }
        }

        if (!packageDefaults)
        {
            styles.AppendChild(TableNormal());
            styles.AppendChild(NoList());
        }

        if (!grid) styles.AppendChild(TableGrid());

        return styles;
    }

    private static BuiltinStyle FromSheet(StyleDefinitionDto style, StyleDefaultsDto defaults)
    {
        var character = style.Type == "character";
        var paragraph = style.Paragraph;
        var run = style.Character;

        return new BuiltinStyle(
            style.Id,
            style.Name,
            Character: character,
            BasedOn: style.BasedOn,
            Next: style.Next,
            Link: style.Link,
            UiPriority: style.UiPriority,
            QFormat: style.QFormat,
            // The model merges `w:hidden` and `w:semiHidden`; Word declares its machinery this way.
            SemiHidden: style.Hidden,
            UnhideWhenUsed: style.Hidden,
            Default: style.Id == (character ? defaults.CharacterStyleId : defaults.ParagraphStyleId),
            SizePt: PointsOf(run?.FontSize),
            Bold: run?.Bold == true,
            Italic: run?.Italic == true,
            Color: run?.Color,
            Underline: run?.Underline == true,
            BeforePt: paragraph?.SpaceBefore,
            AfterPt: paragraph?.SpaceAfter,
            LineFactor: MultipleOf(paragraph?.LineSpacing),
            IndentMm: paragraph?.IndentMm,
            KeepNext: paragraph?.KeepNext == true,
            KeepLines: paragraph?.KeepLines == true,
            ContextualSpacing: paragraph?.ContextualSpacing == true,
            OutlineLevel: paragraph?.OutlineLevel);
    }

    private static string? FirstFontOf(string? stack)
    {
        var first = stack?.Split(',')[0].Trim().Trim('\'', '"');
        return string.IsNullOrEmpty(first) ? null : first;
    }

    private static double? PointsOf(string? size) =>
        size is not null && size.EndsWith("pt", StringComparison.Ordinal) &&
        double.TryParse(size[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var points)
            ? points
            : null;

    private static double? MultipleOf(LineSpacingDto? spacing) =>
        spacing is { Kind: "multiple", Factor: { } factor } ? factor : null;

    public static Style Heading(int level) => Of(BuiltinStyles.Heading(level));

    public static int HeadingLevels => BuiltinStyles.HeadingLevels;

    /// <summary>
    /// In the order the schema fixes, or Word refuses: <c>OpenXmlValidator</c> checks it in the
    /// tests.
    /// </summary>
    public static Style Of(BuiltinStyle style)
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

        if (style.KeepLines) properties.AppendChild(new KeepLines());
        if (SpacingOf(style.BeforePt, style.AfterPt, style.LineFactor) is { } spacing) properties.AppendChild(spacing);

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

        // Also for complex scripts: otherwise Arabic or Hebrew in a heading would come out without
        // bold.
        if (style.Bold)
        {
            properties.AppendChild(new Bold());
            properties.AppendChild(new BoldComplexScript());
        }

        if (style.Italic)
        {
            properties.AppendChild(new Italic());
            properties.AppendChild(new ItalicComplexScript());
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

    private static SpacingBetweenLines? SpacingOf(double? before, double? after, double? factor)
    {
        if (before is null && after is null && factor is null) return null;

        var spacing = new SpacingBetweenLines();
        if (before is { } beforePt) spacing.Before = Twips(beforePt);
        if (after is { } afterPt) spacing.After = Twips(afterPt);
        if (factor is { } lineFactor)
        {
            // 240ths of the natural height, like `ParagraphFormat.ApplyLineHeight`.
            spacing.Line = Invariant((int)Math.Round(lineFactor * 240));
            spacing.LineRule = LineSpacingRuleValues.Auto;
        }

        return spacing;
    }

    /// <summary>
    /// The cell margin is CSS's <c>padding: 4px 8px</c>: 3 pt top and bottom, 6 pt on the sides.
    /// </summary>
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

    /// <summary>Points → half-points (<c>w:sz</c>).</summary>
    private static string HalfPoints(double points) => Invariant((int)Math.Round(points * Unit.HalfPointsPerPoint));

    /// <summary>Points → twips (<c>w:spacing</c>).</summary>
    private static string Twips(double points) => Invariant((int)Math.Round(points * Unit.TwipsPerPoint));

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);
}

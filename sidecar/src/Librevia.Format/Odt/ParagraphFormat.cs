using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>A <c>style:paragraph-properties</c>, each field null when the style says nothing.</summary>
internal sealed record ParagraphFormat(
    string? TextAlign = null,
    double? IndentMm = null,
    double? IndentRightMm = null,
    double? FirstLineMm = null,
    double? SpaceBefore = null,
    double? SpaceAfter = null,
    LineSpacingDto? LineSpacing = null,
    bool? KeepNext = null,
    bool? KeepLines = null,
    bool? WidowControl = null,
    bool? ContextualSpacing = null,
    string? BreakBefore = null,
    string? BreakAfter = null,
    string? Background = null)
{
    public const string Page = "page";
    public const string Column = "column";

    private const string Multiple = "multiple";
    private const string Exact = "exact";
    private const string AtLeast = "atLeast";

    public static readonly ParagraphFormat Empty = new();

    /// <summary>A file that says nothing about line spacing: single.</summary>
    public static readonly LineSpacingDto SingleSpacing = new(Multiple, Factor: 1);

    public ParagraphFormat Over(ParagraphFormat top) => new(
        top.TextAlign ?? TextAlign,
        top.IndentMm ?? IndentMm,
        top.IndentRightMm ?? IndentRightMm,
        top.FirstLineMm ?? FirstLineMm,
        top.SpaceBefore ?? SpaceBefore,
        top.SpaceAfter ?? SpaceAfter,
        top.LineSpacing ?? LineSpacing,
        top.KeepNext ?? KeepNext,
        top.KeepLines ?? KeepLines,
        top.WidowControl ?? WidowControl,
        top.ContextualSpacing ?? ContextualSpacing,
        top.BreakBefore ?? BreakBefore,
        top.BreakAfter ?? BreakAfter,
        top.Background ?? Background);

    public static ParagraphFormat Of(XElement? properties)
    {
        if (properties is null) return Empty;

        string? Fo(string name) => properties.Attr(Ns.Fo + name);
        return new ParagraphFormat(
            TextAlign: AlignOf(Fo("text-align")),
            IndentMm: OdtLength.RoundedMillimetres(Fo("margin-left")),
            IndentRightMm: OdtLength.RoundedMillimetres(Fo("margin-right")),
            FirstLineMm: OdtLength.RoundedMillimetres(Fo("text-indent")),
            SpaceBefore: OdtLength.Points(Fo("margin-top")),
            SpaceAfter: OdtLength.Points(Fo("margin-bottom")),
            LineSpacing: LineSpacingOf(properties),
            KeepNext: KeepOf(Fo("keep-with-next")),
            KeepLines: KeepOf(Fo("keep-together")),
            WidowControl: (Fo("widows") ?? Fo("orphans")) is { } lines ? lines != "0" : null,
            ContextualSpacing: properties.Attr(Ns.Style + "contextual-spacing") is { } contextual ? contextual == "true" : null,
            BreakBefore: BreakOf(Fo("break-before")),
            BreakAfter: BreakOf(Fo("break-after")),
            Background: Fo("background-color") is { } background
                ? background == CharacterFormat.Transparent ? background : CharacterFormat.ColorOf(background)
                : null);
    }

    private static string? AlignOf(string? align) => align switch
    {
        "start" or "left" => "left",
        "end" or "right" => "right",
        "center" => "center",
        "justify" => "justify",
        _ => null,
    };

    private static bool? KeepOf(string? keep) => keep is null ? null : keep == "always";

    /// <summary><c>auto</c> is a statement too: it cancels the style's break.</summary>
    private static string? BreakOf(string? value) => value switch
    {
        Page or Column => value,
        "auto" => "auto",
        _ => null,
    };

    /// <summary>A percentage is a multiple, a length is exact, and the minimum is at least.</summary>
    private static LineSpacingDto? LineSpacingOf(XElement properties)
    {
        if (OdtLength.Points(properties.Attr(Ns.Style + "line-height-at-least")) is { } least)
        {
            return new LineSpacingDto(AtLeast, Points: least);
        }

        var height = properties.Attr(Ns.Fo + "line-height");
        if (OdtLength.Percent(height) is { } factor and > 0) return new LineSpacingDto(Multiple, Factor: Math.Round(factor, 4));
        return OdtLength.Points(height) is { } exact and > 0 ? new LineSpacingDto(Exact, Points: exact) : null;
    }

    public bool PageBreakBefore => BreakBefore == Page;

    /// <summary>Margins only in the paragraph's own measure: a negative indent stays in the margin.</summary>
    public StyleParagraphDto ToDto(int? outlineLevel) => new(
        TextAlign: TextAlign,
        IndentMm: IndentMm,
        IndentRightMm: IndentRightMm,
        FirstLineMm: FirstLineMm,
        SpaceBefore: SpaceBefore,
        SpaceAfter: SpaceAfter,
        LineSpacing: LineSpacing,
        KeepNext: KeepNext,
        KeepLines: KeepLines,
        PageBreakBefore: BreakBefore is null ? null : PageBreakBefore,
        ContextualSpacing: ContextualSpacing,
        OutlineLevel: outlineLevel,
        Background: Background == CharacterFormat.Transparent ? null : Background,
        WidowControl: WidowControl);

    /// <summary>
    /// The CSS <c>line-height</c> the editor wants, as <see cref="BodyReader"/> writes it: a multiple
    /// times the font's natural height, or points.
    /// </summary>
    public static string? CssLineHeight(LineSpacingDto? spacing, string? fontStack)
    {
        if (spacing is null) return null;
        if (spacing.Kind != Multiple) return spacing.Points is { } points ? RunReader.FormatPoints(points) : null;

        var factor = spacing.Factor ?? 1;
        var natural = LineMetrics.Of(fontStack?.Split(',')[0].Trim());
        if (natural is { } known) return Number(Math.Round(factor * known, 4));
        return factor == 1 ? "normal" : Number(Math.Round(factor * NaturalFallback, 4));
    }

    /// <summary>Liberation Serif's, the 1.15 guess <see cref="BodyReader"/> also uses.</summary>
    public const double NaturalFallback = 1.1499;

    private static string Number(double value) =>
        value.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
}

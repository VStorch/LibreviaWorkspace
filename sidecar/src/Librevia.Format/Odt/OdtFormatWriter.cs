using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The inverse of <see cref="ParagraphFormat.Of"/> and <see cref="CharacterFormat.Of"/>: a format
/// record → <c>style:paragraph-properties</c> and <c>style:text-properties</c>, only what the record
/// states. The editor's attributes and marks become those records first.
/// </summary>
internal static class OdtFormatWriter
{
    private const string Multiple = "multiple";
    private const string Exact = "exact";
    private const string AtLeast = "atLeast";
    private const string PointsSuffix = "pt";

    public static XElement Paragraph(ParagraphFormat format)
    {
        var properties = new XElement(Ns.Style + "paragraph-properties");
        void Set(XName name, string? value)
        {
            if (value is not null) properties.SetAttributeValue(name, value);
        }

        Set(Ns.Fo + "text-align", AlignOf(format.TextAlign));
        Set(Ns.Fo + "margin-left", Mm(format.IndentMm));
        Set(Ns.Fo + "margin-right", Mm(format.IndentRightMm));
        Set(Ns.Fo + "text-indent", Mm(format.FirstLineMm));
        Set(Ns.Fo + "margin-top", Pt(format.SpaceBefore));
        Set(Ns.Fo + "margin-bottom", Pt(format.SpaceAfter));
        SetLineSpacing(properties, format.LineSpacing);
        Set(Ns.Fo + "keep-with-next", Keep(format.KeepNext));
        Set(Ns.Fo + "keep-together", Keep(format.KeepLines));
        var lines = format.WidowControl is { } widows ? widows ? "2" : "0" : null;
        Set(Ns.Fo + "widows", lines);
        Set(Ns.Fo + "orphans", lines);
        Set(Ns.Style + "contextual-spacing", Flag(format.ContextualSpacing));
        Set(Ns.Fo + "break-before", format.BreakBefore);
        Set(Ns.Fo + "break-after", format.BreakAfter);
        Set(Ns.Fo + "background-color", format.Background);
        return properties;
    }

    public static XElement Text(CharacterFormat format, OdtFonts fonts)
    {
        var properties = new XElement(Ns.Style + "text-properties");
        void Set(XName name, string? value)
        {
            if (value is not null) properties.SetAttributeValue(name, value);
        }

        if (format.FontFamily is { Length: > 0 } stack)
        {
            if (fonts.NameOf(stack) is { } face) Set(Ns.Style + "font-name", face);
            else Set(Ns.Fo + "font-family", Quoted(fonts.StackOfFamily(stack) ?? stack));
        }

        Set(Ns.Fo + "font-size", Pt(format.FontSizePt));
        Set(Ns.Fo + "font-weight", Choice(format.Bold, "bold", "normal"));
        Set(Ns.Fo + "font-style", Choice(format.Italic, "italic", "normal"));
        SetLine(properties, "underline", format.Underline);
        SetLine(properties, "line-through", format.Strike);
        Set(Ns.Fo + "text-transform", Choice(format.Caps, "uppercase", "none"));
        Set(Ns.Fo + "font-variant", Choice(format.SmallCaps, "small-caps", "normal"));
        Set(Ns.Style + "text-position", PositionOf(format.Position));
        Set(Ns.Fo + "color", format.Color);
        Set(Ns.Fo + "background-color", format.Background);
        return properties;
    }

    /// <summary>A block's attributes, as <see cref="OdtParagraphReader"/> gives them.</summary>
    public static ParagraphFormat OfNode(Node node, string? fontStack) => new(
        TextAlign: Attr.String(node, "textAlign"),
        IndentMm: IndentOf(node),
        IndentRightMm: Attr.Double(node, "indentRightMm"),
        FirstLineMm: Attr.Double(node, "firstLineMm"),
        SpaceBefore: Attr.Double(node, "spaceBefore"),
        SpaceAfter: Attr.Double(node, "spaceAfter"),
        LineSpacing: LineSpacingOf(Attr.Node(node, "lineHeight"), fontStack),
        KeepNext: Optional(node, "keepNext"),
        KeepLines: Optional(node, "keepLines"),
        WidowControl: Optional(node, "widowControl"),
        BreakAfter: Attr.Bool(node, "breakAfter") ? ParagraphFormat.Page
            : Attr.Bool(node, "columnBreakAfter") ? ParagraphFormat.Column : null,
        Background: Attr.String(node, "background"));

    /// <summary>The editor's indent level adds Word's step to the measure, as <c>Docx.ParagraphFormat</c> does.</summary>
    private static double? IndentOf(Node node)
    {
        var measured = Attr.Double(node, "indentMm");
        var level = Attr.Int(node, "indent") ?? 0;
        return level > 0 ? (measured ?? 0) + (level * ListLevels.Mm(Unit.IndentStepTwips)) : measured;
    }

    /// <summary>A style's paragraph format, as <see cref="OdtStyleSheet"/> reads it.</summary>
    public static ParagraphFormat OfStyle(StyleParagraphDto? dto) => dto is null ? ParagraphFormat.Empty : new(
        TextAlign: dto.TextAlign,
        IndentMm: dto.IndentMm,
        IndentRightMm: dto.IndentRightMm,
        FirstLineMm: dto.FirstLineMm,
        SpaceBefore: dto.SpaceBefore,
        SpaceAfter: dto.SpaceAfter,
        LineSpacing: dto.LineSpacing,
        KeepNext: dto.KeepNext,
        KeepLines: dto.KeepLines,
        WidowControl: dto.WidowControl,
        ContextualSpacing: dto.ContextualSpacing,
        BreakBefore: dto.PageBreakBefore is { } pageBreak ? pageBreak ? ParagraphFormat.Page : "auto" : null,
        Background: dto.Background);

    public static CharacterFormat OfStyle(StyleCharacterDto? dto) => dto is null ? CharacterFormat.Empty : new(
        FontFamily: dto.FontFamily,
        FontSizePt: Attr.Points(dto.FontSize),
        Bold: dto.Bold,
        Italic: dto.Italic,
        Underline: dto.Underline,
        Strike: dto.Strike,
        Caps: dto.AllCaps,
        SmallCaps: dto.SmallCaps,
        Position: dto.VerticalAlign,
        Color: dto.Color,
        Background: dto.Highlight);

    /// <summary>
    /// The inverse of <see cref="ParagraphFormat.CssLineHeight"/>: points are exact, a number is a
    /// multiple of the natural height of the font that draws the line.
    /// </summary>
    public static LineSpacingDto? LineSpacingOf(System.Text.Json.Nodes.JsonNode? value, string? fontStack)
    {
        if (value is null) return null;
        double css;
        if (value.GetValueKind() == JsonValueKind.Number)
        {
            css = value.GetValue<double>();
        }
        else if (value.GetValueKind() == JsonValueKind.String)
        {
            var text = value.GetValue<string>().Trim();
            if (text == "normal") return new LineSpacingDto(Multiple, Factor: 1);
            if (text.EndsWith(PointsSuffix, StringComparison.Ordinal)) return Attr.Points(text) is { } points ? new LineSpacingDto(Exact, Points: points) : null;
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out css)) return null;
        }
        else
        {
            return null;
        }

        if (css <= 0) return null;
        var natural = LineMetrics.Of(fontStack?.Split(',')[0].Trim()) ?? ParagraphFormat.NaturalFallback;
        return new LineSpacingDto(Multiple, Factor: Math.Round(css / natural, 4));
    }

    private static void SetLineSpacing(XElement properties, LineSpacingDto? spacing)
    {
        if (spacing is null) return;
        if (spacing.Kind == AtLeast && Pt(spacing.Points) is { } least)
        {
            properties.SetAttributeValue(Ns.Style + "line-height-at-least", least);
        }
        else if (spacing.Kind == Multiple && spacing.Factor is { } factor)
        {
            properties.SetAttributeValue(Ns.Fo + "line-height", Number(factor * 100, "0.##") + "%");
        }
        else if (spacing.Kind == Exact && Pt(spacing.Points) is { } exact)
        {
            properties.SetAttributeValue(Ns.Fo + "line-height", exact);
        }
    }

    private static void SetLine(XElement properties, string kind, bool? on)
    {
        if (on is null) return;
        properties.SetAttributeValue(Ns.Style + $"text-{kind}-style", on.Value ? "solid" : "none");
        if (!on.Value) return;
        if (kind == "underline")
        {
            properties.SetAttributeValue(Ns.Style + "text-underline-width", "auto");
            properties.SetAttributeValue(Ns.Style + "text-underline-color", "font-color");
        }
        else
        {
            properties.SetAttributeValue(Ns.Style + "text-line-through-type", "single");
        }
    }

    private static string? AlignOf(string? align) => align switch
    {
        "left" => "start",
        "right" => "end",
        "center" or "justify" => align,
        _ => null,
    };

    private static string? PositionOf(string? position) => position switch
    {
        CharacterFormat.Superscript => "super 58%",
        CharacterFormat.Subscript => "sub 58%",
        CharacterFormat.Baseline => "0% 100%",
        _ => null,
    };

    private static bool? Optional(Node node, string name) =>
        Attr.Node(node, name) is { } value && value.GetValueKind() is JsonValueKind.True or JsonValueKind.False
            ? value.GetValueKind() == JsonValueKind.True
            : null;

    private static string? Keep(bool? keep) => Choice(keep, "always", "auto");

    private static string? Flag(bool? value) => Choice(value, "true", "false");

    private static string? Choice(bool? value, string on, string off) => value is null ? null : value.Value ? on : off;

    private static string Quoted(string family) => family.Contains(' ') ? $"'{family}'" : family;

    public static string? Mm(double? value) => value is { } mm ? Number(mm, "0.###") + "mm" : null;

    public static string? Pt(double? value) => value is { } points ? Number(points, "0.##") + "pt" : null;

    public static string Number(double value, string format) => value.ToString(format, CultureInfo.InvariantCulture);

    /// <summary>CSS pixels, the editor's measure of images and columns.</summary>
    public static string PixelsToMm(double pixels) => Mm(pixels * Unit.MillimetersPerInch / Unit.PixelsPerInch)!;
}

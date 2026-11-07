using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// A <c>style:text-properties</c>, each field null when the style says nothing. A relative size
/// (<c>130%</c>) waits for the size it is relative to.
/// </summary>
internal sealed record CharacterFormat(
    string? FontFamily = null,
    double? FontSizePt = null,
    double? FontSizeFactor = null,
    bool? Bold = null,
    bool? Italic = null,
    bool? Underline = null,
    bool? Strike = null,
    bool? Caps = null,
    bool? SmallCaps = null,
    string? Position = null,
    string? Color = null,
    string? Background = null)
{
    public const string Superscript = "superscript";
    public const string Subscript = "subscript";
    public const string Baseline = "baseline";
    public const string Transparent = "transparent";

    public static readonly CharacterFormat Empty = new();

    /// <summary><paramref name="top"/> wins where it speaks.</summary>
    public CharacterFormat Over(CharacterFormat top) => new(
        top.FontFamily ?? FontFamily,
        top.FontSizePt ?? (top.FontSizeFactor is { } factor && FontSizePt is { } size ? Math.Round(size * factor, 1) : FontSizePt),
        PendingFactor(top),
        top.Bold ?? Bold,
        top.Italic ?? Italic,
        top.Underline ?? Underline,
        top.Strike ?? Strike,
        top.Caps ?? Caps,
        top.SmallCaps ?? SmallCaps,
        top.Position ?? Position,
        top.Color ?? Color,
        top.Background ?? Background);

    /// <summary>A relative size over a base that does not know its own size yet.</summary>
    private double? PendingFactor(CharacterFormat top)
    {
        if (top.FontSizePt is not null) return null;
        if (top.FontSizeFactor is not { } factor || FontSizePt is not null) return FontSizeFactor;
        return FontSizeFactor is { } own ? own * factor : factor;
    }

    public static CharacterFormat Of(XElement? properties, OdtFonts fonts)
    {
        if (properties is null) return Empty;

        var size = properties.Attr(Ns.Fo + "font-size");
        return new CharacterFormat(
            FontFamily: FontOf(properties, fonts),
            FontSizePt: OdtLength.Points(size),
            FontSizeFactor: OdtLength.Percent(size),
            Bold: WeightOf(properties.Attr(Ns.Fo + "font-weight")),
            Italic: properties.Attr(Ns.Fo + "font-style") is { } style ? style is "italic" or "oblique" : null,
            Underline: LineOf(properties.Attr(Ns.Style + "text-underline-style")),
            Strike: LineOf(properties.Attr(Ns.Style + "text-line-through-style")),
            Caps: properties.Attr(Ns.Fo + "text-transform") is { } transform ? transform == "uppercase" : null,
            SmallCaps: properties.Attr(Ns.Fo + "font-variant") is { } variant ? variant == "small-caps" : null,
            Position: PositionOf(properties.Attr(Ns.Style + "text-position")),
            Color: ColorOf(properties.Attr(Ns.Fo + "color")),
            Background: BackgroundOf(properties.Attr(Ns.Fo + "background-color")));
    }

    private static string? FontOf(XElement properties, OdtFonts fonts)
    {
        if (properties.Attr(Ns.Style + "font-name") is { Length: > 0 } name) return fonts.Stack(name);
        return properties.Attr(Ns.Fo + "font-family") is { Length: > 0 } family ? fonts.StackOfFamily(family) : null;
    }

    /// <summary><c>bold</c> and the numeric weights from 600 up.</summary>
    private static bool? WeightOf(string? weight) => weight switch
    {
        null => null,
        "bold" or "600" or "700" or "800" or "900" => true,
        _ => false,
    };

    private static bool? LineOf(string? style) => style is null ? null : style != "none";

    /// <summary><c>super 58%</c>, <c>33% 58%</c>, <c>-33% 58%</c>, <c>0% 100%</c>.</summary>
    private static string? PositionOf(string? position)
    {
        var first = position?.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        if (first is null) return null;
        if (first == "super") return Superscript;
        if (first == "sub") return Subscript;
        return OdtLength.Percent(first) switch
        {
            > 0 => Superscript,
            < 0 => Subscript,
            _ => Baseline,
        };
    }

    /// <summary>Lowercase <c>#rrggbb</c>; anything else is not a color for the editor.</summary>
    public static string? ColorOf(string? value)
    {
        if (value is not { Length: 7 } || value[0] != '#') return null;
        return value.Skip(1).All(Uri.IsHexDigit) ? value.ToLowerInvariant() : null;
    }

    private static string? BackgroundOf(string? value) =>
        value == Transparent ? Transparent : ColorOf(value);

    public StyleCharacterDto ToDto() => new(
        FontFamily: FontFamily,
        FontSize: FontSizePt is { } size ? RunReader.FormatPoints(size) : null,
        Bold: Bold,
        Italic: Italic,
        Underline: Underline,
        Strike: Strike,
        AllCaps: Caps,
        SmallCaps: SmallCaps,
        VerticalAlign: Position,
        Color: Color,
        Highlight: Background == Transparent ? null : Background);
}

using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// A named style edited field by field: only the attributes of a field the model changed are
/// replaced, so a relative size, a tab stop or a border the panel does not show stays as it was.
/// </summary>
internal static class OdtStyleFields
{
    private sealed record Field<T>(Func<T, object?> Get, Func<T, T> Only, XName[] Names);

    private static readonly Field<ParagraphFormat>[] ParagraphFields =
    [
        new(f => f.TextAlign, f => ParagraphFormat.Empty with { TextAlign = f.TextAlign }, [Ns.Fo + "text-align"]),
        new(f => f.IndentMm, f => ParagraphFormat.Empty with { IndentMm = f.IndentMm }, [Ns.Fo + "margin-left"]),
        new(f => f.IndentRightMm, f => ParagraphFormat.Empty with { IndentRightMm = f.IndentRightMm }, [Ns.Fo + "margin-right"]),
        new(f => f.FirstLineMm, f => ParagraphFormat.Empty with { FirstLineMm = f.FirstLineMm }, [Ns.Fo + "text-indent"]),
        new(f => f.SpaceBefore, f => ParagraphFormat.Empty with { SpaceBefore = f.SpaceBefore }, [Ns.Fo + "margin-top"]),
        new(f => f.SpaceAfter, f => ParagraphFormat.Empty with { SpaceAfter = f.SpaceAfter }, [Ns.Fo + "margin-bottom"]),
        new(f => f.LineSpacing, f => ParagraphFormat.Empty with { LineSpacing = f.LineSpacing },
            [Ns.Fo + "line-height", Ns.Style + "line-height-at-least"]),
        new(f => f.KeepNext, f => ParagraphFormat.Empty with { KeepNext = f.KeepNext }, [Ns.Fo + "keep-with-next"]),
        new(f => f.KeepLines, f => ParagraphFormat.Empty with { KeepLines = f.KeepLines }, [Ns.Fo + "keep-together"]),
        new(f => f.WidowControl, f => ParagraphFormat.Empty with { WidowControl = f.WidowControl }, [Ns.Fo + "widows", Ns.Fo + "orphans"]),
        new(f => f.ContextualSpacing, f => ParagraphFormat.Empty with { ContextualSpacing = f.ContextualSpacing },
            [Ns.Style + "contextual-spacing"]),
        new(f => f.BreakBefore, f => ParagraphFormat.Empty with { BreakBefore = f.BreakBefore }, [Ns.Fo + "break-before"]),
        new(f => f.Background, f => ParagraphFormat.Empty with { Background = f.Background }, [Ns.Fo + "background-color"]),
    ];

    private static readonly Field<CharacterFormat>[] TextFields =
    [
        new(f => f.FontFamily, f => CharacterFormat.Empty with { FontFamily = f.FontFamily },
            [Ns.Style + "font-name", Ns.Fo + "font-family"]),
        new(f => f.FontSizePt, f => CharacterFormat.Empty with { FontSizePt = f.FontSizePt }, [Ns.Fo + "font-size"]),
        new(f => f.Bold, f => CharacterFormat.Empty with { Bold = f.Bold }, [Ns.Fo + "font-weight"]),
        new(f => f.Italic, f => CharacterFormat.Empty with { Italic = f.Italic }, [Ns.Fo + "font-style"]),
        new(f => f.Underline, f => CharacterFormat.Empty with { Underline = f.Underline },
            [Ns.Style + "text-underline-style", Ns.Style + "text-underline-width", Ns.Style + "text-underline-color"]),
        new(f => f.Strike, f => CharacterFormat.Empty with { Strike = f.Strike },
            [Ns.Style + "text-line-through-style", Ns.Style + "text-line-through-type"]),
        new(f => f.Caps, f => CharacterFormat.Empty with { Caps = f.Caps }, [Ns.Fo + "text-transform"]),
        new(f => f.SmallCaps, f => CharacterFormat.Empty with { SmallCaps = f.SmallCaps }, [Ns.Fo + "font-variant"]),
        new(f => f.Position, f => CharacterFormat.Empty with { Position = f.Position }, [Ns.Style + "text-position"]),
        new(f => f.Color, f => CharacterFormat.Empty with { Color = f.Color }, [Ns.Fo + "color"]),
        new(f => f.Background, f => CharacterFormat.Empty with { Background = f.Background }, [Ns.Fo + "background-color"]),
    ];

    public static void Paragraph(XElement properties, ParagraphFormat known, ParagraphFormat wanted) =>
        Apply(properties, ParagraphFields, known, wanted, OdtFormatWriter.Paragraph);

    public static void Text(XElement properties, CharacterFormat known, CharacterFormat wanted, OdtFonts fonts) =>
        Apply(properties, TextFields, known, wanted, format => OdtFormatWriter.Text(format, fonts));

    private static void Apply<T>(XElement properties, Field<T>[] fields, T known, T wanted, Func<T, XElement> write)
    {
        foreach (var field in fields)
        {
            if (Equals(field.Get(known), field.Get(wanted))) continue;
            foreach (var name in field.Names) properties.Attribute(name)?.Remove();
            foreach (var attribute in write(field.Only(wanted)).Attributes()) properties.SetAttributeValue(attribute.Name, attribute.Value);
        }
    }
}

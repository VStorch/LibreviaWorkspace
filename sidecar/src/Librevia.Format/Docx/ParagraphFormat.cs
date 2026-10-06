using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

using static Librevia.Format.Docx.InvariantText;

namespace Librevia.Format.Docx;

/// <summary>
/// The edited paragraph's <c>w:pPr</c>: it **starts from the original** and only overrides what the
/// model represents. Built from scratch, it would take style, spacing, background and the
/// <c>w:sectPr</c> closing the section; what the editor does not know passes through intact.
/// </summary>
/// <remarks>
/// On a loose body paragraph the node carries only direct formatting, and it is the whole truth:
/// what left it was cleared and leaves the file. Lists, cells and old drafts (<paramref
/// name="flatten"/>) arrive flattened, and there absence says nothing. Either way, what only
/// repeats the style is not written where the original did not declare it; <paramref
/// name="styles"/> says what the style is worth for line spacing and per-level indent.
/// </remarks>
internal sealed class ParagraphFormat(
    Inventory inventory,
    HeadingStyles headings,
    StyleResolver styles,
    bool flatten = false,
    bool revisions = true)
{
    /// <summary>Or <c>null</c> when there is nothing to say.</summary>
    /// <param name="list"><c>null</c> when the caller does not know the numbering: the original
    /// <c>w:numPr</c> stays.</param>
    public ParagraphProperties? Build(Node node, ParagraphWriter.ListPlacement? list, Paragraph? original)
    {
        var properties = original?.ParagraphProperties?.CloneNode(true) as ParagraphProperties
                         ?? new ParagraphProperties();

        ApplyStyle(properties, node);

        // The node is the whole truth only on a loose paragraph read without flattening.
        var direct = !flatten &&
                     list is { List: null } &&
                     original?.ParagraphProperties?.NumberingProperties is null;
        if (direct) ClearWhatWasCleared(properties, node);
        var style = styles.StyleParagraphOf(properties);

        ApplyAlignment(properties, node);
        ApplyIndentation(properties, node);
        ApplyExplicitZeros(properties, node);
        ApplySpacing(properties, node);
        ApplyShading(properties, node);
        ApplyKeepNext(properties, node, direct);
        ApplyKeepLines(properties, node, direct);
        ApplyWidowControl(properties, node, direct);
        ApplyMark(properties, node);
        if (revisions) ApplyMarkRevision(properties, node);

        DropWhatRepeatsTheStyle(properties, original?.ParagraphProperties, style);
        DropMarkThatRepeatsTheStyle(properties, original?.ParagraphProperties, styles.Resolve(properties).Run);

        // Numbering comes from the context. Without a wrapper (cell, box) it means "unknown", and
        // the original `w:numPr` stays; with an empty wrapper it means "stopped being an item".
        if (list is { } placement)
        {
            properties.NumberingProperties = placement.List is not { } context
                ? null
                : new NumberingProperties(
                    new NumberingLevelReference { Val = context.Level },
                    new NumberingId { Val = context.NumberingId });
        }

        return properties.HasChildren ? properties : null;
    }

    /// <summary>
    /// The <c>styleId</c> the model carries, not one recomputed from the level, which would replace
    /// LibreOffice's <c>Ttulo1</c> with a <c>Heading1</c> the document does not define. The level
    /// only rules on a paragraph that became a heading here, through the <c>heading N</c> style id
    /// (<see cref="HeadingStyles"/>). And only an id **this** package defines counts.
    /// </summary>
    private void ApplyStyle(ParagraphProperties properties, Node node)
    {
        var declared = Attr.String(node, "styleId");
        var level = node.Type == "heading" ? Attr.Int(node, "level") : null;

        if (level is not null)
        {
            var keep = declared is not null &&
                       headings.Defines(declared) &&
                       (BodyReader.HeadingLevelOfStyle(declared) == level || headings.LevelByName(declared) == level);
            var heading = keep ? declared! : headings.IdFor(level.Value);
            properties.ParagraphStyleId = new ParagraphStyleId { Val = heading };
            NoteUndefined(heading);
            return;
        }

        // A heading that stopped being one no longer points to the heading style, by id and by
        // name.
        if (declared is null ||
            BodyReader.HeadingLevelOfStyle(declared) is not null ||
            headings.LevelByName(declared) is not null)
        {
            if (declared is not null) properties.ParagraphStyleId = null;
            return;
        }

        var id = headings.IdForDeclared(declared);
        properties.ParagraphStyleId = new ParagraphStyleId { Val = id };
        NoteUndefined(id);
    }

    /// <summary>
    /// Word draws Normal for a style the package does not define: the loss goes to the inventory.
    /// </summary>
    private void NoteUndefined(string id)
    {
        if (!headings.Defines(id)) inventory.NoteLoss($"estilo \"{id}\", que o documento não define");
    }

    private static void ApplyAlignment(ParagraphProperties properties, Node node)
    {
        if (JustificationOf(Attr.String(node, "textAlign")) is not { } justification) return;
        properties.Justification = justification;
    }

    /// <summary>
    /// Or <c>null</c> without a requested alignment. Public because block images use it too.
    /// </summary>
    public static Justification? JustificationOf(string? align) => align switch
    {
        null => null,
        "center" => new Justification { Val = JustificationValues.Center },
        "right" => new Justification { Val = JustificationValues.Right },
        "justify" => new Justification { Val = JustificationValues.Both },
        _ => new Justification { Val = JustificationValues.Left },
    };

    /// <remarks>
    /// The file's measure, plus the <c>Ctrl+]</c> level. No indent on the model is an **assertion**
    /// (the <c>Ctrl+[</c> all the way). A negative indent never reaches the editor, so it stays.
    /// </remarks>
    private void ApplyIndentation(ParagraphProperties properties, Node node)
    {
        // Without a direct measure, the paragraph's indent is the style's, and the level adds to
        // it.
        var level = Attr.Int(node, "indent") ?? 0;
        var measured = Attr.MmToTwips(Attr.Double(node, "indentMm"))
                       ?? (level > 0 ? StyleTwips(styles.StyleParagraphOf(properties).Indentation?.Left) : 0);
        var left = measured + (level * Unit.IndentStepTwips);
        var right = Attr.MmToTwips(Attr.Double(node, "indentRightMm")) ?? 0;
        var firstLine = Attr.MmToTwips(Attr.Double(node, "firstLineMm")) ?? 0;

        if (left <= 0 && right <= 0 && firstLine == 0)
        {
            ClearIndentation(properties.Indentation);
            return;
        }

        var indentation = properties.Indentation;
        if (indentation is null)
        {
            indentation = new Indentation();
            properties.Indentation = indentation;
        }

        // `Twips(...)`: a ternary on `string` would write `w:right=""`, which Word refuses.
        indentation.Left = Twips(left > 0 ? left : null);
        indentation.Right = Twips(right > 0 ? right : null);

        // A single attribute, with the sign deciding which.
        indentation.FirstLine = Twips(firstLine > 0 ? firstLine : null);
        indentation.Hanging = Twips(firstLine < 0 ? -firstLine : null);
    }

    /// <summary>
    /// Only what the model represents; nobody can have cleared borders, tabs or <c>w:sectPr</c>.
    /// The node's "transparent" <c>w:shd</c> is the original's colorless one, and it stays.
    /// </summary>
    private static void ClearWhatWasCleared(ParagraphProperties properties, Node node)
    {
        bool Absent(string name) => Attr.Node(node, name) is null;

        // The reader always carries `w:pStyle`: without it on the node, it was removed.
        if (node.Type != "heading" && Absent("styleId")) properties.ParagraphStyleId = null;

        if (Absent("textAlign")) properties.Justification = null;

        if (properties.Indentation is { } indentation)
        {
            if (Absent("indentMm") && (Attr.Int(node, "indent") ?? 0) == 0)
            {
                indentation.Left = null;
                indentation.Start = null;
            }

            if (Absent("indentRightMm"))
            {
                indentation.Right = null;
                indentation.End = null;
            }

            if (Absent("firstLineMm"))
            {
                indentation.FirstLine = null;
                indentation.Hanging = null;
            }

            if (!indentation.HasAttributes) properties.Indentation = null;
        }

        if (properties.SpacingBetweenLines is { } spacing)
        {
            if (Absent("spaceBefore")) spacing.Before = null;
            if (Absent("spaceAfter")) spacing.After = null;
            if (Absent("lineHeight"))
            {
                spacing.Line = null;
                spacing.LineRule = null;
            }

            if (!spacing.HasAttributes) properties.SpacingBetweenLines = null;
        }

        if (Absent("background")) properties.Shading = null;
        if (Absent("keepNext")) properties.KeepNext = null;
        if (Absent("keepLines")) properties.KeepLines = null;
        if (Absent("widowControl")) properties.WidowControl = null;

        if (properties.ParagraphMarkRunProperties is { } mark)
        {
            if (Absent("fontFamily") && mark.GetFirstChild<RunFonts>() is { } fonts)
            {
                fonts.Ascii = null;
                fonts.HighAnsi = null;
                if (!fonts.HasAttributes) fonts.Remove();
            }

            if (Absent("fontSize")) mark.GetFirstChild<FontSize>()?.Remove();
        }
    }

    /// <summary>
    /// Only what the original did **not** declare; what was declared was the author's choice.
    /// </summary>
    private static void DropWhatRepeatsTheStyle(
        ParagraphProperties properties,
        ParagraphProperties? original,
        ParagraphProperties style)
    {
        if (original?.Justification is null && properties.Justification is { } jc &&
            SameValue(jc.Val?.InnerText, style.Justification?.Val?.InnerText ?? "left"))
        {
            properties.Justification = null;
        }

        if (properties.Indentation is { } indentation)
        {
            var was = original?.Indentation;
            var from = style.Indentation;
            if (was?.Left is null && SameTwips(indentation.Left, from?.Left)) indentation.Left = null;
            if (was?.Right is null && SameTwips(indentation.Right, from?.Right)) indentation.Right = null;
            if (was?.FirstLine is null && was?.Hanging is null &&
                SameTwips(indentation.FirstLine, from?.FirstLine) &&
                SameTwips(indentation.Hanging, from?.Hanging))
            {
                indentation.FirstLine = null;
                indentation.Hanging = null;
            }

            if (!indentation.HasAttributes) properties.Indentation = null;
        }

        if (properties.SpacingBetweenLines is { } spacing)
        {
            var was = original?.SpacingBetweenLines;
            var from = style.SpacingBetweenLines;
            if (was?.Before is null && SameTwips(spacing.Before, from?.Before)) spacing.Before = null;
            if (was?.After is null && SameTwips(spacing.After, from?.After)) spacing.After = null;
            if (was?.Line is null && spacing.Line is not null &&
                SameValue(spacing.Line.Value, from?.Line?.Value ?? "240") &&
                SameValue(spacing.LineRule?.InnerText ?? "auto", from?.LineRule?.InnerText ?? "auto"))
            {
                spacing.Line = null;
                spacing.LineRule = null;
            }

            if (!spacing.HasAttributes) properties.SpacingBetweenLines = null;
        }

        if (original?.Shading is null && properties.Shading is { } shading &&
            SameValue(shading.Fill?.Value, style.Shading?.Fill?.Value))
        {
            properties.Shading = null;
        }

        if (original?.KeepNext is null && properties.KeepNext is { } keep &&
            RunReader.IsOn(keep) == RunReader.IsOn(style.KeepNext))
        {
            properties.KeepNext = null;
        }

        if (original?.KeepLines is null && properties.KeepLines is { } lines &&
            RunReader.IsOn(lines) == RunReader.IsOn(style.KeepLines))
        {
            properties.KeepLines = null;
        }

        // A silent style keeps widow control: it is Word's default.
        if (original?.WidowControl is null && properties.WidowControl is { } widow &&
            RunReader.IsOn(widow) == (style.WidowControl is null || RunReader.IsOn(style.WidowControl)))
        {
            properties.WidowControl = null;
        }

        if (properties.ParagraphMarkRunProperties is { HasChildren: false }) properties.ParagraphMarkRunProperties = null;
    }

    private static void DropMarkThatRepeatsTheStyle(
        ParagraphProperties properties,
        ParagraphProperties? original,
        RunProperties style)
    {
        if (properties.ParagraphMarkRunProperties is not { } mark) return;
        var was = original?.ParagraphMarkRunProperties;

        if (was?.GetFirstChild<RunFonts>() is null && mark.GetFirstChild<RunFonts>() is { } fonts &&
            SameValue(fonts.Ascii?.Value, style.RunFonts?.Ascii?.Value))
        {
            fonts.Remove();
        }

        if (was?.GetFirstChild<FontSize>() is null && mark.GetFirstChild<FontSize>() is { } size &&
            SameValue(size.Val?.Value, style.FontSize?.Val?.Value))
        {
            size.Remove();
        }

        if (!mark.HasChildren) properties.ParagraphMarkRunProperties = null;
    }

    /// <summary>Word draws a missing measure as zero.</summary>
    private static bool SameTwips(StringValue? a, StringValue? b) =>
        a is not null && string.Equals(a.Value ?? "0", b?.Value ?? "0", StringComparison.Ordinal);

    private static bool SameValue(string? a, string? b) =>
        a is not null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private static int StyleTwips(StringValue? measure) =>
        int.TryParse(measure?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips) && twips > 0
            ? twips
            : 0;

    /// <summary>
    /// Zero on the block undoes the style's indent: without the attribute, Word would indent again.
    /// </summary>
    private static void ApplyExplicitZeros(ParagraphProperties properties, Node node)
    {
        var left = Attr.Double(node, "indentMm") is 0 && (Attr.Int(node, "indent") ?? 0) == 0;
        var right = Attr.Double(node, "indentRightMm") is 0;
        var firstLine = Attr.Double(node, "firstLineMm") is 0;
        if (!left && !right && !firstLine) return;

        var indentation = properties.Indentation ??= new Indentation();
        if (left && indentation.Left is null) indentation.Left = "0";
        if (right && indentation.Right is null) indentation.Right = "0";
        if (firstLine && indentation.FirstLine is null && indentation.Hanging is null) indentation.FirstLine = "0";
    }

    private static void ClearIndentation(Indentation? indentation)
    {
        if (indentation is null) return;

        // Explicit zero: deleting the attribute would bring the style's indent back.
        if (IsPositive(indentation.Left)) indentation.Left = "0";
        if (IsPositive(indentation.Right)) indentation.Right = "0";

        // `w:firstLine` and `w:hanging`: here zeroing means removing.
        if (IsPositive(indentation.FirstLine)) indentation.FirstLine = null;
        if (IsPositive(indentation.Hanging)) indentation.Hanging = null;
    }

    /// <summary>
    /// Something that is not a whole twip (<c>0.5in</c>) returns false and stays where it is.
    /// </summary>
    private static bool IsPositive(StringValue? measure) =>
        int.TryParse(measure?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
        && twips > 0;

    private void ApplySpacing(ParagraphProperties properties, Node node)
    {
        var before = Attr.Double(node, "spaceBefore");
        var after = Attr.Double(node, "spaceAfter");
        var lineHeight = Attr.String(node, "lineHeight");
        if (before is null && after is null && lineHeight is null) return;

        var spacing = properties.SpacingBetweenLines;
        if (spacing is null)
        {
            spacing = new SpacingBetweenLines();
            properties.SpacingBetweenLines = spacing;
        }

        // Twentieths: exact to a tenth of a point, the reader's precision.
        if (before is not null) spacing.Before = Invariant((int)Math.Round(before.Value * Unit.TwipsPerPoint));
        if (after is not null) spacing.After = Invariant((int)Math.Round(after.Value * Unit.TwipsPerPoint));
        if (lineHeight is not null) ApplyLineHeight(spacing, lineHeight, Attr.String(node, "fontFamily") ?? MarkFontOf(properties));
    }

    /// <summary>
    /// The inverse of <c>BodyReader.LineHeightOf</c>: 240ths of the paragraph font's natural
    /// height. <c>normal</c> is not converted: it is the file's silence.
    /// </summary>
    private void ApplyLineHeight(SpacingBetweenLines spacing, string value, string? fontStack)
    {
        if (value.Equals("normal", StringComparison.OrdinalIgnoreCase)) return;

        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            if (Attr.Points(value) is not { } points || points <= 0) return;

            spacing.Line = Invariant((int)Math.Round(points * Unit.TwipsPerPoint));

            // "At least", not "exactly": `exact` clips what does not fit.
            if (spacing.LineRule is null || spacing.LineRule.Value == LineSpacingRuleValues.Auto)
            {
                spacing.LineRule = LineSpacingRuleValues.AtLeast;
            }

            return;
        }

        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var css) || css <= 0)
        {
            inventory.NoteLoss($"entrelinha \"{value}\"");
            return;
        }

        var natural = LineMetrics.Of(FirstFont(fontStack)) ?? 1.1499;
        var factor = css / natural;
        if (factor is <= 0.5 or >= 4)
        {
            inventory.NoteLoss($"entrelinha \"{value}\"");
            return;
        }

        spacing.Line = Invariant((int)Math.Round(factor * 240));
        spacing.LineRule = LineSpacingRuleValues.Auto;
    }

    /// <summary>
    /// The style's with the original mark on top: the same one the reader multiplied with.
    /// </summary>
    private string? MarkFontOf(ParagraphProperties properties) =>
        styles.ResolveMark(styles.Resolve(properties).Run, properties).RunFonts?.Ascii?.Value;

    private void ApplyShading(ParagraphProperties properties, Node node)
    {
        if (Attr.String(node, "background") is not { } background) return;

        // The colorless `w:shd` is already in the clone.
        if (background.Equals("transparent", StringComparison.OrdinalIgnoreCase)) return;

        if (ColorValue.Hex(background) is not { } fill)
        {
            inventory.NoteLoss($"cor de fundo \"{background}\"");
            return;
        }

        properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
    }

    /// <remarks>
    /// Turning off means deleting the element, only when it was on: <c>w:val="false"</c> exists to
    /// turn off the style's.
    /// </remarks>
    private static void ApplyKeepNext(ParagraphProperties properties, Node node, bool direct)
    {
        if (Attr.Bool(node, "keepNext"))
        {
            properties.KeepNext = new KeepNext();
            return;
        }

        // On a direct node, `false` undoes the style's, and only an explicit off says so.
        if (direct && Attr.Node(node, "keepNext") is not null)
        {
            properties.KeepNext = new KeepNext { Val = false };
            return;
        }

        if (RunReader.IsOn(properties.KeepNext)) properties.KeepNext = null;
    }

    /// <remarks>
    /// In Word it is **on** when nothing says otherwise: on is only written when it undoes an off.
    /// </remarks>
    private static void ApplyWidowControl(ParagraphProperties properties, Node node, bool direct)
    {
        if (Attr.Node(node, "widowControl") is not null && !Attr.Bool(node, "widowControl"))
        {
            properties.WidowControl = new WidowControl { Val = false };
            return;
        }

        if (Attr.Bool(node, "widowControl") && direct)
        {
            properties.WidowControl = new WidowControl();
            return;
        }

        if (properties.WidowControl is { } widow && !RunReader.IsOn(widow)) properties.WidowControl = new WidowControl();
    }

    /// <remarks>The same rule as <see cref="ApplyKeepNext"/>, for `w:keepLines`.</remarks>
    private static void ApplyKeepLines(ParagraphProperties properties, Node node, bool direct)
    {
        if (Attr.Bool(node, "keepLines"))
        {
            properties.KeepLines = new KeepLines();
            return;
        }

        if (direct && Attr.Node(node, "keepLines") is not null)
        {
            properties.KeepLines = new KeepLines { Val = false };
            return;
        }

        if (RunReader.IsOn(properties.KeepLines)) properties.KeepLines = null;
    }

    /// <summary>
    /// The paragraph mark font (<c>w:pPr/w:rPr</c>), which measures the line; the rest of the mark
    /// stays.
    /// </summary>
    private void ApplyMark(ParagraphProperties properties, Node node)
    {
        var font = Attr.String(node, "fontFamily");
        var size = Attr.String(node, "fontSize");
        if (font is null && size is null) return;

        var mark = properties.ParagraphMarkRunProperties;
        if (mark is null)
        {
            mark = new ParagraphMarkRunProperties();
            properties.ParagraphMarkRunProperties = mark;
        }

        if (FirstFont(font) is { } family)
        {
            var fonts = mark.GetFirstChild<RunFonts>();
            if (fonts is null) PutInOrder(mark, new RunFonts { Ascii = family, HighAnsi = family });
            else
            {
                fonts.Ascii = family;
                fonts.HighAnsi = family;
            }
        }

        if (size is null) return;

        if (Attr.Points(size) is { } points && points > 0)
        {
            var halfPoints = Invariant((int)Math.Round(points * Unit.HalfPointsPerPoint));
            var declared = mark.GetFirstChild<FontSize>();
            if (declared is null) PutInOrder(mark, new FontSize { Val = halfPoints });
            else declared.Val = halfPoints;
        }
        else
        {
            inventory.NoteLoss($"tamanho de fonte \"{size}\"");
        }
    }

    /// <c>markRevision</c> ↔ <c>w:pPr/w:rPr/w:ins|w:del</c>; the file's stays when it is the same.
    private static void ApplyMarkRevision(ParagraphProperties properties, Node node)
    {
        var wanted = Attr.Node(node, "markRevision");
        var mark = properties.ParagraphMarkRunProperties;
        if (mark is null)
        {
            if (wanted is null) return;
            mark = new ParagraphMarkRunProperties();
            properties.ParagraphMarkRunProperties = mark;
        }

        Revisions.ApplyBlock(mark, wanted, created => PutInOrder(mark, created));
        if (!mark.HasChildren) properties.ParagraphMarkRunProperties = null;
    }

    /// <summary>
    /// OOXML is a sequence: <c>w:sz</c> after <c>w:u</c> invalidates the document. The paragraph
    /// mark is the only one the SDK does not expose typed.
    /// </summary>
    private static readonly string[] MarkOrder =
    [
        "ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
        "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
        "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish",
        "oMath", "rPrChange",
    ];

    private static void PutInOrder(ParagraphMarkRunProperties mark, OpenXmlElement child)
    {
        var rank = Array.IndexOf(MarkOrder, child.LocalName);

        // Unknown elements go to the end, where vendor extensions declare themselves.
        var next = mark.ChildElements.FirstOrDefault(existing =>
        {
            var other = Array.IndexOf(MarkOrder, existing.LocalName);
            return other < 0 || other > rank;
        });

        if (next is null) mark.AppendChild(child);
        else mark.InsertBefore(child, next);
    }

    /// <c>w:rFonts</c> stores a name, and the reader delivers a stack.
    internal static string? FirstFont(string? stack)
    {
        if (stack is null) return null;
        var first = stack.Split(',')[0].Trim().Trim('\'', '"');
        return first.Length > 0 ? first : null;
    }

    /// <c>StringValue?</c>: <c>null</c> deletes the attribute. A ternary on <c>string</c> would
    /// write <c>w:ind w:right=""</c>, outside the schema.
    private static StringValue? Twips(int? value) =>
        value is null ? null : new StringValue(Invariant(value.Value));
}

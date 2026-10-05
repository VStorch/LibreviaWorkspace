using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// The original's styles are reread by <see cref="StyleReader"/> and compared record by record:
/// without differences, <c>word/styles.xml</c> goes back byte for byte. What changed is altered
/// **in place**, field by field. A new style goes at the end with <c>w:customStyle</c>; none is
/// deleted, because a part the editor does not open may point to it; and only custom styles get
/// renamed.
/// </summary>
internal static class StyleWriter
{
    /// <summary>The copy Word 2010 writes alongside, which we do not rewrite.</summary>
    internal const string StaleEffects = "estilos com efeitos do Word 2010 (stylesWithEffects.xml)";

    /// <returns>Whether anything was written.</returns>
    /// <param name="additionsOnly">
    /// A flattened draft: only styles the package lacks. Version 2 styles were invented in the
    /// migration (<c>LEGACY_STYLES</c>), and writing them would rewrite the real ones in Times; a
    /// change to an existing one is declared as a loss.
    /// </param>
    public static bool Apply(
        MainDocumentPart part,
        StyleSheetDto? wanted,
        Inventory inventory,
        HashSet<string> touched,
        bool additionsOnly = false)
    {
        if (wanted is null) return false;

        var current = StyleReader.Read(part);
        if (current.Defaults == wanted.Defaults && SameStyles(current, wanted)) return false;

        var definitions = part.StyleDefinitionsPart ?? part.AddNewPart<StyleDefinitionsPart>();
        var styles = definitions.Styles ??= new Styles();
        var changed = false;

        var skipped = false;

        if (additionsOnly)
        {
            skipped = current.Defaults != wanted.Defaults;
        }
        else if (current.Defaults.Paragraph != wanted.Defaults.Paragraph)
        {
            var defaults = styles.DocDefaults ??= new DocDefaults();
            var holder = defaults.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault();
            var properties = holder.ParagraphPropertiesBaseStyle ??= new ParagraphPropertiesBaseStyle();
            Paragraph(properties, current.Defaults.Paragraph, wanted.Defaults.Paragraph);
            changed = true;
        }

        if (!additionsOnly && current.Defaults.Character != wanted.Defaults.Character)
        {
            var defaults = styles.DocDefaults ??= new DocDefaults();
            var holder = defaults.RunPropertiesDefault ??= new RunPropertiesDefault();
            var properties = holder.RunPropertiesBaseStyle ??= new RunPropertiesBaseStyle();
            Character(properties, current.Defaults.Character, wanted.Defaults.Character);
            changed = true;
        }

        foreach (var (id, definition) in wanted.Styles)
        {
            if (current.Styles.TryGetValue(id, out var before))
            {
                if (before == definition) continue;
                if (additionsOnly)
                {
                    skipped = true;
                    continue;
                }

                if (StyleOf(styles, id) is not { } element) continue;

                Modify(styles, element, before, definition, current, inventory);
                changed = true;
                continue;
            }

            styles.AppendChild(Created(definition));
            changed = true;
        }

        if (skipped) inventory.NoteLoss("mudança nos estilos existentes feita num rascunho de versão antiga");
        if (!changed) return false;

        styles.Save();
        touched.Add(definitions.Uri.OriginalString.TrimStart('/'));

        // Word reads `styles.xml`; only very old readers prefer the copy, and the warning says so.
        if (part.StylesWithEffectsPart is not null) inventory.NoteInvisible(StaleEffects);
        return true;
    }

    private static bool SameStyles(StyleSheetDto current, StyleSheetDto wanted) =>
        wanted.Styles.All(entry => current.Styles.TryGetValue(entry.Key, out var before) && before == entry.Value);

    /// <summary>The first <c>w:style</c> with the id, as the reader and the resolver do.</summary>
    private static Style? StyleOf(Styles styles, string id) =>
        styles.Elements<Style>().FirstOrDefault(style => style.StyleId?.Value == id);

    private static void Modify(
        Styles styles,
        Style element,
        StyleDefinitionDto before,
        StyleDefinitionDto after,
        StyleSheetDto current,
        Inventory inventory)
    {
        // A builtin style's name (`heading 1`) does not change; the UI does not offer it, and
        // whoever asked is told.
        if (before.Name != after.Name)
        {
            if (before.Custom) element.StyleName = new StyleName { Val = after.Name };
            else inventory.NoteLoss($"novo nome do estilo embutido \"{before.Name}\"");
        }
        if (before.BasedOn != after.BasedOn) element.BasedOn = after.BasedOn is null ? null : new BasedOn { Val = after.BasedOn };
        if (before.Next != after.Next)
        {
            element.NextParagraphStyle = after.Next is null ? null : new NextParagraphStyle { Val = after.Next };
        }

        if (before.Paragraph != after.Paragraph)
        {
            var properties = element.StyleParagraphProperties ??= new StyleParagraphProperties();
            Paragraph(properties, before.Paragraph ?? new StyleParagraphDto(), after.Paragraph ?? new StyleParagraphDto());
            if (!properties.HasChildren) element.StyleParagraphProperties = null;
        }

        if (before.Character == after.Character) return;

        var run = element.StyleRunProperties ??= new StyleRunProperties();
        Character(run, before.Character ?? new StyleCharacterDto(), after.Character ?? new StyleCharacterDto());
        if (!run.HasChildren) element.StyleRunProperties = null;

        // `w:link` is the other half of the style: with a different font, the two would diverge.
        if (after.Link is { } link && current.Styles.TryGetValue(link, out var linked) &&
            linked.Type == "character" && StyleOf(styles, link) is { } partner)
        {
            var partnerRun = partner.StyleRunProperties ??= new StyleRunProperties();
            Character(partnerRun, linked.Character ?? new StyleCharacterDto(), after.Character ?? new StyleCharacterDto());
            if (!partnerRun.HasChildren) partner.StyleRunProperties = null;
        }
    }

    private static Style Created(StyleDefinitionDto definition)
    {
        var style = new Style
        {
            Type = definition.Type == "character" ? StyleValues.Character : StyleValues.Paragraph,
            StyleId = definition.Id,
            // "1", as Word writes it, not the SDK's "true".
            CustomStyle = new OnOffValue { InnerText = "1" },
        };

        // In schema order: name, basedOn, next, link, priority, qFormat, properties.
        style.AppendChild(new StyleName { Val = definition.Name });
        if (definition.BasedOn is not null) style.AppendChild(new BasedOn { Val = definition.BasedOn });
        if (definition.Next is not null) style.AppendChild(new NextParagraphStyle { Val = definition.Next });
        if (definition.Link is not null) style.AppendChild(new LinkedStyle { Val = definition.Link });
        if (definition.UiPriority is { } priority) style.AppendChild(new UIPriority { Val = priority });
        if (definition.QFormat) style.AppendChild(new PrimaryStyle());

        if (definition.Paragraph is { } paragraph && definition.Type != "character")
        {
            var properties = new StyleParagraphProperties();
            Paragraph(properties, new StyleParagraphDto(), paragraph);
            if (properties.HasChildren) style.AppendChild(properties);
        }

        if (definition.Character is { } character)
        {
            var run = new StyleRunProperties();
            Character(run, new StyleCharacterDto(), character);
            if (run.HasChildren) style.AppendChild(run);
        }

        return style;
    }


    /// <summary>Only the fields that changed; null removes.</summary>
    private static void Paragraph(OpenXmlElement properties, StyleParagraphDto before, StyleParagraphDto after)
    {
        if (before.TextAlign != after.TextAlign)
        {
            Put(properties, "jc", after.TextAlign is null ? null : ParagraphFormat.JustificationOf(after.TextAlign), PPrOrder);
        }

        if (before.IndentMm != after.IndentMm || before.IndentRightMm != after.IndentRightMm ||
            before.FirstLineMm != after.FirstLineMm)
        {
            var indentation = properties.GetFirstChild<Indentation>() ?? new Indentation();
            // Character measures (`*Chars`) beat twips in Word: they go.
            if (before.IndentMm != after.IndentMm)
            {
                indentation.Left = Twips(after.IndentMm);
                indentation.Start = null;
                indentation.LeftChars = null;
                indentation.StartCharacters = null;
            }

            if (before.IndentRightMm != after.IndentRightMm)
            {
                indentation.Right = Twips(after.IndentRightMm);
                indentation.End = null;
                indentation.RightChars = null;
                indentation.EndCharacters = null;
            }

            if (before.FirstLineMm != after.FirstLineMm)
            {
                indentation.FirstLineChars = null;
                indentation.HangingChars = null;
                indentation.FirstLine = after.FirstLineMm is >= 0 ? Twips(after.FirstLineMm) : null;
                indentation.Hanging = after.FirstLineMm is < 0 ? Twips(-after.FirstLineMm) : null;
            }

            Put(properties, "ind", indentation.HasAttributes ? indentation : null, PPrOrder);
        }

        if (before.SpaceBefore != after.SpaceBefore || before.SpaceAfter != after.SpaceAfter ||
            before.LineSpacing != after.LineSpacing)
        {
            var spacing = properties.GetFirstChild<SpacingBetweenLines>() ?? new SpacingBetweenLines();
            // Lines and the HTML automatic beat twips, as for indents.
            if (before.SpaceBefore != after.SpaceBefore)
            {
                spacing.Before = PointsToTwips(after.SpaceBefore);
                spacing.BeforeLines = null;
                spacing.BeforeAutoSpacing = null;
            }

            if (before.SpaceAfter != after.SpaceAfter)
            {
                spacing.After = PointsToTwips(after.SpaceAfter);
                spacing.AfterLines = null;
                spacing.AfterAutoSpacing = null;
            }
            if (before.LineSpacing != after.LineSpacing) LineSpacing(spacing, after.LineSpacing);
            Put(properties, "spacing", spacing.HasAttributes ? spacing : null, PPrOrder);
        }

        if (before.KeepNext != after.KeepNext) Put(properties, "keepNext", Toggle<KeepNext>(after.KeepNext), PPrOrder);
        if (before.KeepLines != after.KeepLines) Put(properties, "keepLines", Toggle<KeepLines>(after.KeepLines), PPrOrder);
        if (before.WidowControl != after.WidowControl)
        {
            Put(properties, "widowControl", Toggle<WidowControl>(after.WidowControl), PPrOrder);
        }
        if (before.PageBreakBefore != after.PageBreakBefore)
        {
            Put(properties, "pageBreakBefore", Toggle<PageBreakBefore>(after.PageBreakBefore), PPrOrder);
        }

        if (before.ContextualSpacing != after.ContextualSpacing)
        {
            Put(properties, "contextualSpacing", Toggle<ContextualSpacing>(after.ContextualSpacing), PPrOrder);
        }

        if (before.OutlineLevel != after.OutlineLevel)
        {
            Put(properties, "outlineLvl", after.OutlineLevel is { } level ? new OutlineLevel { Val = level } : null, PPrOrder);
        }

        if (before.Background != after.Background)
        {
            Put(properties, "shd", ShadingOf(after.Background), PPrOrder);
        }
    }

    private static void LineSpacing(SpacingBetweenLines spacing, LineSpacingDto? value)
    {
        switch (value)
        {
            case null:
                spacing.Line = null;
                spacing.LineRule = null;
                break;
            case { Kind: "multiple", Factor: { } factor }:
                spacing.Line = Invariant(Rounded(factor * 240));
                spacing.LineRule = LineSpacingRuleValues.Auto;
                break;
            case { Points: { } points }:
                spacing.Line = Invariant(Rounded(points * Unit.TwipsPerPoint));
                spacing.LineRule = value.Kind == "exact" ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.AtLeast;
                break;
        }
    }


    private static void Character(OpenXmlElement properties, StyleCharacterDto before, StyleCharacterDto after)
    {
        if (before.FontFamily != after.FontFamily)
        {
            var fonts = properties.GetFirstChild<RunFonts>() ?? new RunFonts();
            var family = ParagraphFormat.FirstFont(after.FontFamily);
            fonts.Ascii = family;
            fonts.HighAnsi = family;
            // A theme font would beat the declared one.
            fonts.AsciiTheme = null;
            fonts.HighAnsiTheme = null;
            Put(properties, "rFonts", fonts.HasAttributes ? fonts : null, RPrOrder);
        }

        if (before.FontSize != after.FontSize)
        {
            var size = Attr.Points(after.FontSize) is { } points and > 0
                ? new FontSize { Val = Invariant(Rounded(points * Unit.HalfPointsPerPoint)) }
                : null;
            Put(properties, "sz", size, RPrOrder);
        }

        if (before.Bold != after.Bold) Put(properties, "b", Toggle<Bold>(after.Bold), RPrOrder);
        if (before.Italic != after.Italic) Put(properties, "i", Toggle<Italic>(after.Italic), RPrOrder);
        if (before.Strike != after.Strike) Put(properties, "strike", Toggle<Strike>(after.Strike), RPrOrder);
        if (before.AllCaps != after.AllCaps) Put(properties, "caps", Toggle<Caps>(after.AllCaps), RPrOrder);
        if (before.SmallCaps != after.SmallCaps) Put(properties, "smallCaps", Toggle<SmallCaps>(after.SmallCaps), RPrOrder);

        if (before.Underline != after.Underline)
        {
            var line = after.Underline switch
            {
                null => null,
                true => new Underline { Val = UnderlineValues.Single },
                false => new Underline { Val = UnderlineValues.None },
            };
            Put(properties, "u", line, RPrOrder);
        }

        if (before.VerticalAlign != after.VerticalAlign)
        {
            var position = after.VerticalAlign switch
            {
                "super" => new VerticalTextAlignment { Val = VerticalPositionValues.Superscript },
                "sub" => new VerticalTextAlignment { Val = VerticalPositionValues.Subscript },
                _ => null,
            };
            Put(properties, "vertAlign", position, RPrOrder);
        }

        if (before.Color != after.Color)
        {
            var color = ColorValue.Hex(after.Color) is { } hex ? new Color { Val = hex } : null;
            Put(properties, "color", color, RPrOrder);
        }

        // Goes back as `w:shd`, which holds any color; a named `w:highlight` goes so it does not
        // beat it.
        if (before.Highlight != after.Highlight)
        {
            properties.GetFirstChild<Highlight>()?.Remove();
            Put(properties, "shd", ShadingOf(after.Highlight), RPrOrder);
        }
    }


    /// <summary>
    /// On without <c>w:val</c>, off with <c>w:val="0"</c>, silence without the element.
    /// </summary>
    private static T? Toggle<T>(bool? value) where T : OnOffType, new() => value switch
    {
        null => null,
        true => new T(),
        false => new T { Val = false },
    };

    private static Shading? ShadingOf(string? color) =>
        ColorValue.Hex(color) is { } fill
            ? new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill }
            : null;

    private static StringValue? Twips(double? millimeters) =>
        millimeters is { } mm ? new StringValue(Invariant(Attr.MmToTwips(mm))) : null;

    private static StringValue? PointsToTwips(double? points) =>
        points is { } value ? new StringValue(Invariant(Rounded(value * Unit.TwipsPerPoint))) : null;

    /// <summary>
    /// Away from zero, like <see cref="Attr.MmToTwips(double)"/>: 10.25 pt is 21 half-points.
    /// </summary>
    private static int Rounded(double value) => (int)Math.Round(value, MidpointRounding.AwayFromZero);

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Null removes; at its schema slot; a child already in place stays where it is.
    /// </summary>
    private static void Put(OpenXmlElement parent, string name, OpenXmlElement? child, string[] order)
    {
        var old = parent.ChildElements.FirstOrDefault(element => element.LocalName == name);
        if (child is not null && ReferenceEquals(old, child)) return;
        old?.Remove();
        if (child is null) return;

        var rank = Array.IndexOf(order, name);
        var next = parent.ChildElements.FirstOrDefault(element =>
        {
            var other = Array.IndexOf(order, element.LocalName);
            return other < 0 || other > rank;
        });

        if (next is null) parent.AppendChild(child);
        else parent.InsertBefore(child, next);
    }

    private static readonly string[] PPrOrder =
    [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
        "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
        "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    ];

    /// <summary>The same as the paragraph mark's.</summary>
    private static readonly string[] RPrOrder =
    [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
        "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
        "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
        "em", "lang", "eastAsianLayout", "specVanish", "oMath",
    ];
}

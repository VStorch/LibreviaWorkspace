using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Grava de volta os estilos que o modelo mudou ou criou.
/// </summary>
/// <remarks>
/// A mesma regra da gravação cirúrgica, um andar acima: o que não mudou não é
/// tocado. As definições do original são relidas pelo <see cref="StyleReader"/> —
/// a mesma leitura que produziu o modelo — e comparadas estilo a estilo pela
/// igualdade dos registros, que é a impressão digital da definição. Sem nenhuma
/// diferença, `word/styles.xml` não entra nas partes tocadas e volta byte a byte.
///
/// O estilo que mudou é alterado **no lugar**, campo a campo: só o que o modelo
/// representa e que de fato mudou é sobreposto ou removido. O resto do `w:style`
/// — `w:rsid`, `w:tblStylePr`, borda, tabulação, idioma, fonte de tema do campo
/// que ninguém mexeu — fica como estava, assim como `w:latentStyles` e os estilos
/// de tabela e de numeração, que o modelo nem enxerga.
///
/// Estilo novo entra no fim, com `w:customStyle`. Nenhum estilo é excluído: um
/// estilo que sumiu do modelo pode ainda ser apontado por uma parte que o editor
/// não abre (cabeçalho, nota, comentário). E só o personalizado muda de nome — o
/// nome do embutido (`heading 1`) é como o Word o reconhece.
/// </remarks>
internal static class StyleWriter
{
    /// <summary>A cópia com efeitos que o Word 2010 grava ao lado, e que não regravamos.</summary>
    internal const string StaleEffects = "estilos com efeitos do Word 2010 (stylesWithEffects.xml)";

    /// <returns>Se alguma coisa foi gravada.</returns>
    public static bool Apply(MainDocumentPart part, StyleSheetDto? wanted, Inventory inventory, HashSet<string> touched)
    {
        if (wanted is null) return false;

        var current = StyleReader.Read(part);
        if (current.Defaults == wanted.Defaults && SameStyles(current, wanted)) return false;

        var definitions = part.StyleDefinitionsPart ?? part.AddNewPart<StyleDefinitionsPart>();
        var styles = definitions.Styles ??= new Styles();
        var changed = false;

        if (current.Defaults.Paragraph != wanted.Defaults.Paragraph)
        {
            var defaults = styles.DocDefaults ??= new DocDefaults();
            var holder = defaults.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault();
            var properties = holder.ParagraphPropertiesBaseStyle ??= new ParagraphPropertiesBaseStyle();
            Paragraph(properties, current.Defaults.Paragraph, wanted.Defaults.Paragraph);
            changed = true;
        }

        if (current.Defaults.Character != wanted.Defaults.Character)
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
                if (StyleOf(styles, id) is not { } element) continue;

                Modify(styles, element, before, definition, current);
                changed = true;
                continue;
            }

            styles.AppendChild(Created(definition));
            changed = true;
        }

        if (!changed) return false;

        styles.Save();
        touched.Add(definitions.Uri.OriginalString.TrimStart('/'));

        // A cópia antiga fica com as definições de antes. O Word lê `styles.xml`;
        // só leitor muito antigo prefere a outra, e é disso que o aviso fala.
        if (part.StylesWithEffectsPart is not null) inventory.NoteInvisible(StaleEffects);
        return true;
    }

    private static bool SameStyles(StyleSheetDto current, StyleSheetDto wanted) =>
        wanted.Styles.All(entry => current.Styles.TryGetValue(entry.Key, out var before) && before == entry.Value);

    /// <summary>O primeiro `w:style` com o id — o mesmo que o leitor e o resolvedor usam.</summary>
    private static Style? StyleOf(Styles styles, string id) =>
        styles.Elements<Style>().FirstOrDefault(style => style.StyleId?.Value == id);

    private static void Modify(
        Styles styles,
        Style element,
        StyleDefinitionDto before,
        StyleDefinitionDto after,
        StyleSheetDto current)
    {
        if (before.Custom && before.Name != after.Name) element.StyleName = new StyleName { Val = after.Name };
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

        // O estilo de caractere ligado (`w:link`) é a outra metade do mesmo
        // estilo: o Word o aplica ao trecho quando se escolhe o estilo de
        // parágrafo numa parte da linha. Com fonte diferente, as duas metades
        // divergiriam.
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
            // "1", e não o "true" que o SDK escreve: é a forma que o Word grava.
            CustomStyle = new OnOffValue { InnerText = "1" },
        };

        // A ordem é a do esquema: nome, herança, seguinte, ligado, prioridade,
        // qFormat, e só então as propriedades.
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

    // --- parágrafo ------------------------------------------------------------

    /// <summary>Só os campos que mudaram; nulo remove o que o arquivo dizia.</summary>
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
            if (before.IndentMm != after.IndentMm)
            {
                indentation.Left = Twips(after.IndentMm);
                indentation.Start = null;
            }

            if (before.IndentRightMm != after.IndentRightMm)
            {
                indentation.Right = Twips(after.IndentRightMm);
                indentation.End = null;
            }

            if (before.FirstLineMm != after.FirstLineMm)
            {
                indentation.FirstLine = after.FirstLineMm is >= 0 ? Twips(after.FirstLineMm) : null;
                indentation.Hanging = after.FirstLineMm is < 0 ? Twips(-after.FirstLineMm) : null;
            }

            Put(properties, "ind", indentation.HasAttributes ? indentation : null, PPrOrder);
        }

        if (before.SpaceBefore != after.SpaceBefore || before.SpaceAfter != after.SpaceAfter ||
            before.LineSpacing != after.LineSpacing)
        {
            var spacing = properties.GetFirstChild<SpacingBetweenLines>() ?? new SpacingBetweenLines();
            if (before.SpaceBefore != after.SpaceBefore) spacing.Before = PointsToTwips(after.SpaceBefore);
            if (before.SpaceAfter != after.SpaceAfter) spacing.After = PointsToTwips(after.SpaceAfter);
            if (before.LineSpacing != after.LineSpacing) LineSpacing(spacing, after.LineSpacing);
            Put(properties, "spacing", spacing.HasAttributes ? spacing : null, PPrOrder);
        }

        if (before.KeepNext != after.KeepNext) Put(properties, "keepNext", Toggle<KeepNext>(after.KeepNext), PPrOrder);
        if (before.KeepLines != after.KeepLines) Put(properties, "keepLines", Toggle<KeepLines>(after.KeepLines), PPrOrder);
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
                spacing.Line = Invariant((int)Math.Round(factor * 240));
                spacing.LineRule = LineSpacingRuleValues.Auto;
                break;
            case { Points: { } points }:
                spacing.Line = Invariant((int)Math.Round(points * 20));
                spacing.LineRule = value.Kind == "exact" ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.AtLeast;
                break;
        }
    }

    // --- caractere ------------------------------------------------------------

    private static void Character(OpenXmlElement properties, StyleCharacterDto before, StyleCharacterDto after)
    {
        if (before.FontFamily != after.FontFamily)
        {
            var fonts = properties.GetFirstChild<RunFonts>() ?? new RunFonts();
            var family = ParagraphFormat.FirstFont(after.FontFamily);
            fonts.Ascii = family;
            fonts.HighAnsi = family;
            // A fonte de tema venceria a declarada: o Word a lê primeiro.
            fonts.AsciiTheme = null;
            fonts.HighAnsiTheme = null;
            Put(properties, "rFonts", fonts.HasAttributes ? fonts : null, RPrOrder);
        }

        if (before.FontSize != after.FontSize)
        {
            var size = Attr.Points(after.FontSize) is { } points and > 0
                ? new FontSize { Val = Invariant((int)Math.Round(points * 2)) }
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

        // O realce é lido do `w:highlight` nomeado ou do `w:shd`; volta como
        // `w:shd`, que guarda qualquer cor, e o nomeado sai para não vencê-lo.
        if (before.Highlight != after.Highlight)
        {
            properties.GetFirstChild<Highlight>()?.Remove();
            Put(properties, "shd", ShadingOf(after.Highlight), RPrOrder);
        }
    }

    // --- peças ----------------------------------------------------------------

    /// <summary>Ligado sem `w:val`, desligado com `w:val="0"`, silêncio sem elemento.</summary>
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
        points is { } value ? new StringValue(Invariant((int)Math.Round(value * 20))) : null;

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Troca o filho de nome <paramref name="name"/> por <paramref name="child"/>
    /// (nulo remove), na posição que o esquema manda. O próprio filho já no lugar
    /// — alterado por dentro — fica onde está.
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

    /// <summary>A sequência do `w:pPr`, como o esquema a fixa.</summary>
    private static readonly string[] PPrOrder =
    [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
        "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
        "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    ];

    /// <summary>A sequência do `w:rPr` — a mesma da marca de parágrafo.</summary>
    private static readonly string[] RPrOrder =
    [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
        "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
        "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
        "em", "lang", "eastAsianLayout", "specVanish", "oMath",
    ];
}

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O <c>w:pPr</c> do parágrafo editado: **parte-se do original** e só se sobrepõe o
/// que o modelo representa. Montado do zero, levaria estilo, espaçamento, fundo e
/// o <c>w:sectPr</c> que fecha a seção; o que o editor não conhece atravessa intacto.
/// </summary>
/// <remarks>
/// No parágrafo solto do corpo o nó traz só o direto, e é a verdade inteira: o que
/// sumiu dele foi limpo e sai do arquivo. Lista, célula e rascunho antigo
/// (<paramref name="flatten"/>) chegam achatados, e ali ausência não diz nada. Em
/// qualquer caso, o que só repete o estilo não é gravado onde o original não o
/// declarava; <paramref name="styles"/> diz o que o estilo vale para a entrelinha
/// e o recuo por nível.
/// </remarks>
internal sealed class ParagraphFormat(
    Inventory inventory,
    HeadingStyles headings,
    StyleResolver styles,
    bool flatten = false,
    bool revisions = true)
{
    private const int TwipsPerIndentLevel = 720;

    /// <summary>Ou <c>null</c> quando não há nada a dizer.</summary>
    /// <param name="list"><c>null</c> quando quem grava não sabe a numeração: o <c>w:numPr</c> do original fica.</param>
    public ParagraphProperties? Build(Node node, ParagraphWriter.ListPlacement? list, Paragraph? original)
    {
        var properties = original?.ParagraphProperties?.CloneNode(true) as ParagraphProperties
                         ?? new ParagraphProperties();

        ApplyStyle(properties, node);

        // O nó é a verdade inteira só no parágrafo solto lido sem achatar.
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

        // A numeração vem do contexto. Sem embrulho (célula, caixa) é "não sei", e
        // o `w:numPr` do original fica; com embrulho vazio é "deixou de ser item".
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
    /// O <c>styleId</c> que o modelo carrega, e não um recalculado do nível, que
    /// trocaria o <c>Ttulo1</c> do LibreOffice por um <c>Heading1</c> que o documento não
    /// define. O nível só manda no parágrafo que virou título aqui, pelo id do estilo
    /// <c>heading N</c> (<see cref="HeadingStyles"/>). E só vale o id que **este** pacote define.
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

        // Título que deixou de ser título não aponta mais o estilo de título, pelo id e pelo nome.
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

    /// <summary>O Word desenha o Normal para estilo que o pacote não define: a perda vai ao inventário.</summary>
    private void NoteUndefined(string id)
    {
        if (!headings.Defines(id)) inventory.NoteLoss($"estilo \"{id}\", que o documento não define");
    }

    private static void ApplyAlignment(ParagraphProperties properties, Node node)
    {
        if (JustificationOf(Attr.String(node, "textAlign")) is not { } justification) return;
        properties.Justification = justification;
    }

    /// <summary>Ou <c>null</c> sem alinhamento pedido. Público porque a imagem em bloco também o usa.</summary>
    public static Justification? JustificationOf(string? align) => align switch
    {
        null => null,
        "center" => new Justification { Val = JustificationValues.Center },
        "right" => new Justification { Val = JustificationValues.Right },
        "justify" => new Justification { Val = JustificationValues.Both },
        _ => new Justification { Val = JustificationValues.Left },
    };

    /// <remarks>
    /// A medida do arquivo, e o nível de <c>Ctrl+]</c> somado a ela. Sem recuo no modelo
    /// é **afirmação** (o <c>Ctrl+[</c> até o fim). O recuo negativo nunca chega ao
    /// editor, e por isso fica.
    /// </remarks>
    private void ApplyIndentation(ParagraphProperties properties, Node node)
    {
        // Sem medida direta, o recuo que o parágrafo tem é o do estilo, e o nível soma a ele.
        var level = Attr.Int(node, "indent") ?? 0;
        var measured = Attr.MmToTwips(Attr.Double(node, "indentMm"))
                       ?? (level > 0 ? StyleTwips(styles.StyleParagraphOf(properties).Indentation?.Left) : 0);
        var left = measured + (level * TwipsPerIndentLevel);
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

        // `Twips(...)`: o ternário sobre `string` gravaria `w:right=""`, que o Word recusa.
        indentation.Left = Twips(left > 0 ? left : null);
        indentation.Right = Twips(right > 0 ? right : null);

        // Um atributo só, com o sinal decidindo qual.
        indentation.FirstLine = Twips(firstLine > 0 ? firstLine : null);
        indentation.Hanging = Twips(firstLine < 0 ? -firstLine : null);
    }

    /// <summary>
    /// Só o que o modelo representa; bordas, tabulações e <c>w:sectPr</c> ninguém pode
    /// ter limpado. O <c>w:shd</c> "transparente" do nó é o sem cor do original, e fica.
    /// </summary>
    private static void ClearWhatWasCleared(ParagraphProperties properties, Node node)
    {
        bool Absent(string name) => Attr.Node(node, name) is null;

        // O leitor sempre leva o `w:pStyle`: sem ele no nó, foi tirado.
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

    /// <summary>Só o que o original **não** declarava; o declarado foi escolha de quem escreveu.</summary>
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

        // O estilo que cala controla viúvas: é o padrão do Word.
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

    /// <summary>O Word desenha a medida ausente como zero.</summary>
    private static bool SameTwips(StringValue? a, StringValue? b) =>
        a is not null && string.Equals(a.Value ?? "0", b?.Value ?? "0", StringComparison.Ordinal);

    private static bool SameValue(string? a, string? b) =>
        a is not null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private static int StyleTwips(StringValue? measure) =>
        int.TryParse(measure?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips) && twips > 0
            ? twips
            : 0;

    /// <summary>Zero no bloco desfaz o recuo do estilo: sem o atributo, o Word voltaria a recuar.</summary>
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

        // Zero explícito: apagar o atributo traria o recuo do estilo de volta.
        if (IsPositive(indentation.Left)) indentation.Left = "0";
        if (IsPositive(indentation.Right)) indentation.Right = "0";

        // `w:firstLine` e `w:hanging`: aqui zerar é remover.
        if (IsPositive(indentation.FirstLine)) indentation.FirstLine = null;
        if (IsPositive(indentation.Hanging)) indentation.Hanging = null;
    }

    /// <summary>O que não é twip inteiro (<c>0.5in</c>) devolve falso e fica onde está.</summary>
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

        // Vinte avos: exato até o décimo de ponto, a precisão do leitor.
        if (before is not null) spacing.Before = Invariant((int)Math.Round(before.Value * 20));
        if (after is not null) spacing.After = Invariant((int)Math.Round(after.Value * 20));
        if (lineHeight is not null) ApplyLineHeight(spacing, lineHeight, Attr.String(node, "fontFamily") ?? MarkFontOf(properties));
    }

    /// <summary>
    /// O inverso de <c>BodyReader.LineHeightOf</c>: 240-avos da altura natural da
    /// fonte do parágrafo. <c>normal</c> não é convertido: é o silêncio do arquivo.
    /// </summary>
    private void ApplyLineHeight(SpacingBetweenLines spacing, string value, string? fontStack)
    {
        if (value.Equals("normal", StringComparison.OrdinalIgnoreCase)) return;

        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            if (Attr.Points(value) is not { } points || points <= 0) return;

            spacing.Line = Invariant((int)Math.Round(points * 20));

            // "Pelo menos", e não "exatamente": `exact` corta o que não cabe.
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

    /// <summary>A do estilo com a marca original por cima: a mesma com que o leitor multiplicou.</summary>
    private string? MarkFontOf(ParagraphProperties properties) =>
        styles.ResolveMark(styles.Resolve(properties).Run, properties).RunFonts?.Ascii?.Value;

    private void ApplyShading(ParagraphProperties properties, Node node)
    {
        if (Attr.String(node, "background") is not { } background) return;

        // O `w:shd` sem cor já está no clone.
        if (background.Equals("transparent", StringComparison.OrdinalIgnoreCase)) return;

        if (ColorValue.Hex(background) is not { } fill)
        {
            inventory.NoteLoss($"cor de fundo \"{background}\"");
            return;
        }

        properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
    }

    /// <remarks>
    /// Desligar é apagar o elemento, só quando estava ligado: <c>w:val="false"</c> existe
    /// para desligar o do estilo.
    /// </remarks>
    private static void ApplyKeepNext(ParagraphProperties properties, Node node, bool direct)
    {
        if (Attr.Bool(node, "keepNext"))
        {
            properties.KeepNext = new KeepNext();
            return;
        }

        // No nó direto, `false` desfaz o do estilo, e só o desligado explícito o diz.
        if (direct && Attr.Node(node, "keepNext") is not null)
        {
            properties.KeepNext = new KeepNext { Val = false };
            return;
        }

        if (RunReader.IsOn(properties.KeepNext)) properties.KeepNext = null;
    }

    /// <remarks>
    /// No Word vale **ligado** quando nada diz: o ligado só é escrito quando desfaz um
    /// desligado.
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

    /// <remarks>A mesma regra do <see cref="ApplyKeepNext"/>, para `w:keepLines`.</remarks>
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

    /// <summary>A fonte da marca de parágrafo (<c>w:pPr/w:rPr</c>), que mede a linha; o resto da marca fica.</summary>
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
            var halfPoints = Invariant((int)Math.Round(points * 2));
            var declared = mark.GetFirstChild<FontSize>();
            if (declared is null) PutInOrder(mark, new FontSize { Val = halfPoints });
            else declared.Val = halfPoints;
        }
        else
        {
            inventory.NoteLoss($"tamanho de fonte \"{size}\"");
        }
    }

    /// <summary><c>markRevision</c> ↔ <c>w:pPr/w:rPr/w:ins|w:del</c>; a do arquivo fica quando é a mesma.</summary>
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
    /// O OOXML é sequência: <c>w:sz</c> depois de <c>w:u</c> invalida o documento. A marca de
    /// parágrafo é a única que o SDK não expõe tipada.
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

        // O desconhecido vai para o fim, onde extensão de fornecedor se declara.
        var next = mark.ChildElements.FirstOrDefault(existing =>
        {
            var other = Array.IndexOf(MarkOrder, existing.LocalName);
            return other < 0 || other > rank;
        });

        if (next is null) mark.AppendChild(child);
        else mark.InsertBefore(child, next);
    }

    /// <summary>O <c>w:rFonts</c> guarda um nome, e o leitor entrega uma pilha.</summary>
    internal static string? FirstFont(string? stack)
    {
        if (stack is null) return null;
        var first = stack.Split(',')[0].Trim().Trim('\'', '"');
        return first.Length > 0 ? first : null;
    }

    /// <summary>Formata, sem a vírgula de nenhuma região: serve a twips, meios-pontos e 240-avos.</summary>
    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// <c>StringValue?</c>: o <c>null</c> apaga o atributo. Um ternário sobre <c>string</c>
    /// gravaria <c>w:ind w:right=""</c>, fora do esquema.
    /// </summary>
    private static StringValue? Twips(int? value) =>
        value is null ? null : new StringValue(Invariant(value.Value));
}

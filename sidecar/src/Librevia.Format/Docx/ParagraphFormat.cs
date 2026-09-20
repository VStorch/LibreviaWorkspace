using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O `w:pPr` de um parágrafo editado: o original, com o que o modelo diz por
/// cima.
/// </summary>
/// <remarks>
/// Antes daqui o `w:pPr` era montado do zero a cada bloco reescrito, e com ele
/// iam embora o estilo do parágrafo, o espaçamento, a entrelinha, o fundo, o
/// "manter com o próximo", a fonte da marca de parágrafo e — pior de tudo — o
/// `w:sectPr` que fecha uma seção. O leitor lê todas essas coisas e o editor as
/// mostra na tela; o escritor as apagava em silêncio ao gravar. Corrigir uma
/// vírgula num título custava o título.
///
/// A regra é a mesma da gravação cirúrgica, um nível abaixo: **parte-se do XML
/// original** e só se sobrepõe o que o modelo de fato representa. O que o editor
/// não conhece — `w:pBdr`, `w:framePr`, `w:tabs`, `w:sectPr`, a marca de
/// parágrafo inteira — atravessa a edição intacto porque ninguém o reescreve.
/// </remarks>
internal sealed class ParagraphFormat(Inventory inventory)
{
    private const int TwipsPerIndentLevel = 720;

    /// <summary>
    /// O `w:pPr` a gravar, ou <c>null</c> quando não há nada a dizer.
    /// </summary>
    /// <param name="list">
    /// O que quem grava sabe sobre a numeração: <c>null</c> quando não sabe nada,
    /// e aí o `w:numPr` do original fica onde está.
    /// </param>
    /// <param name="original">O parágrafo como estava no arquivo, quando existe.</param>
    public ParagraphProperties? Build(Node node, ParagraphWriter.ListPlacement? list, Paragraph? original)
    {
        var properties = original?.ParagraphProperties?.CloneNode(true) as ParagraphProperties
                         ?? new ParagraphProperties();

        ApplyStyle(properties, node);
        ApplyAlignment(properties, node);
        ApplyIndentation(properties, node);
        ApplySpacing(properties, node);
        ApplyShading(properties, node);
        ApplyKeepNext(properties, node);
        ApplyMark(properties, node);

        // A numeração vem do contexto, e não do nó: no arquivo a lista são
        // parágrafos irmãos apontando o mesmo `w:numId`.
        //
        // Quem grava dentro de uma célula ou de uma caixa de texto não passa
        // contexto nenhum — `null` —, e isso não quer dizer "não é lista": quer
        // dizer "não sei". Aí o `w:numPr` do original fica de pé. Com embrulho,
        // a afirmação é do corpo: vazio por dentro é "deixou de ser item", e o
        // apontamento sai, senão o parágrafo continua numerado depois de a pessoa
        // ter tirado a lista.
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
    /// O estilo do parágrafo: o que o modelo carrega, e não um nome inventado.
    /// </summary>
    /// <remarks>
    /// O `styleId` viaja no modelo justamente para isto. Recalculá-lo a partir do
    /// nível do título trocava `Ttulo1` — como o LibreOffice o grava — por
    /// `Heading1`, que o documento não define: o título perdia a formatação
    /// inteira ao ser editado.
    ///
    /// O nível só manda quando o estilo que o modelo traz não é de título: é o
    /// caso do parágrafo que a pessoa transformou em título aqui dentro, e aí não
    /// há estilo original a respeitar.
    /// </remarks>
    private static void ApplyStyle(ParagraphProperties properties, Node node)
    {
        var declared = Attr.String(node, "styleId");
        var level = node.Type == "heading" ? Attr.Int(node, "level") : null;

        if (level is not null)
        {
            var keep = declared is not null && BodyReader.HeadingLevelOfStyle(declared) == level;
            properties.ParagraphStyleId = new ParagraphStyleId { Val = keep ? declared : "Heading" + level };
            return;
        }

        // Título que deixou de ser título não pode continuar apontando o estilo
        // de título: na tela virou parágrafo, e no arquivo continuaria barra
        // vermelha.
        if (declared is null || BodyReader.HeadingLevelOfStyle(declared) is not null)
        {
            if (declared is not null) properties.ParagraphStyleId = null;
            return;
        }

        properties.ParagraphStyleId = new ParagraphStyleId { Val = declared };
    }

    private static void ApplyAlignment(ParagraphProperties properties, Node node)
    {
        if (Attr.String(node, "textAlign") is not { } align) return;

        properties.Justification = new Justification
        {
            Val = align switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                "justify" => JustificationValues.Both,
                _ => JustificationValues.Left,
            },
        };
    }

    /// <remarks>
    /// A medida do arquivo primeiro, e o nível do editor por cima: são as duas
    /// origens do recuo, e o nível existe porque `Ctrl+]` trabalha em passos.
    /// Somá-los é o que faz recuar um parágrafo importado acrescentar um passo ao
    /// recuo que ele já tinha, em vez de apagá-lo.
    ///
    /// Modelo sem recuo nenhum é **afirmação**, e não silêncio: é o que volta de
    /// quem apertou `Ctrl+[` até o fim. Enquanto este método saía calado nesse
    /// caso, a diminuição não chegava ao arquivo e o recuo reaparecia ao reabrir
    /// o documento.
    ///
    /// Zera só o que o leitor de fato mostra na tela. Um recuo **negativo** — a
    /// primeira linha que sai para fora da margem — nunca chega ao editor, porque
    /// o leitor só emite a medida positiva; o zero que volta do modelo não está
    /// falando dele, e por isso ele fica.
    /// </remarks>
    private static void ApplyIndentation(ParagraphProperties properties, Node node)
    {
        var left = (Attr.MmToTwips(Attr.Double(node, "indentMm")) ?? 0)
                   + ((Attr.Int(node, "indent") ?? 0) * TwipsPerIndentLevel);
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

        // `Twips(...)`, e não o ternário direto: `x > 0 ? Invariant(x) : null`
        // tem tipo `string`, e o `null` de um `string` vira `StringValue(null)` na
        // conversão implícita — o SDK então grava `w:right=""` em vez de omitir o
        // atributo. O LibreOffice tolera; o Word recusa o documento inteiro.
        indentation.Left = Twips(left > 0 ? left : null);
        indentation.Right = Twips(right > 0 ? right : null);

        // Um só atributo, com o sinal decidindo qual: o Word grava `w:firstLine`
        // e `w:hanging` como dois, e declarar os dois deixaria o arquivo dizendo
        // duas coisas sobre a mesma linha.
        indentation.FirstLine = Twips(firstLine > 0 ? firstLine : null);
        indentation.Hanging = Twips(firstLine < 0 ? -firstLine : null);
    }

    /// <summary>O recuo positivo do arquivo, zerado — o negativo fica de pé.</summary>
    private static void ClearIndentation(Indentation? indentation)
    {
        if (indentation is null) return;

        // Zero explícito, e não a ausência do atributo: o estilo do parágrafo
        // também declara recuo, e apagar o atributo traria o do estilo de volta.
        if (IsPositive(indentation.Left)) indentation.Left = "0";
        if (IsPositive(indentation.Right)) indentation.Right = "0";

        // `w:firstLine` e `w:hanging` empurram e puxam a mesma linha, e nenhum é
        // o padrão do outro: aqui zerar é remover.
        if (IsPositive(indentation.FirstLine)) indentation.FirstLine = null;
        if (IsPositive(indentation.Hanging)) indentation.Hanging = null;
    }

    /// <summary>
    /// A medida do arquivo é positiva?
    /// </summary>
    /// <remarks>
    /// O que não é número inteiro de twips — o OOXML moderno aceita `0.5in` —
    /// devolve falso e fica onde está: não sabemos comparar, e mexer no que não
    /// se entende é como a medida do autor se perde.
    /// </remarks>
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

        // Ponto → twip: vinte avos, e é exato nas duas direções até o décimo de
        // ponto, que é a precisão com que o leitor entrega a medida.
        if (before is not null) spacing.Before = Invariant((int)Math.Round(before.Value * 20));
        if (after is not null) spacing.After = Invariant((int)Math.Round(after.Value * 20));
        if (lineHeight is not null) ApplyLineHeight(spacing, lineHeight, Attr.String(node, "fontFamily"));
    }

    /// <summary>
    /// A entrelinha do CSS de volta para `w:line`.
    /// </summary>
    /// <remarks>
    /// O caminho de ida está em <c>BodyReader.LineHeightOf</c>, e este é o
    /// inverso dele. O múltiplo do OOXML vem em 240-avos e é medido sobre a
    /// **altura natural da fonte**, não sobre o tamanho dela — por isso a conta
    /// precisa saber qual fonte o parágrafo usa, que é a mesma informação que a
    /// leitura usou para produzir o número.
    ///
    /// `normal` não é convertido: quer dizer "a altura que a fonte pedir", que é
    /// justamente o silêncio do arquivo. Sobrescrever o `w:line` original com um
    /// palpite de 240 mudaria a paginação de um parágrafo que ninguém tocou na
    /// entrelinha.
    /// </remarks>
    private void ApplyLineHeight(SpacingBetweenLines spacing, string value, string? fontStack)
    {
        if (value.Equals("normal", StringComparison.OrdinalIgnoreCase)) return;

        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            if (Attr.Points(value) is not { } points || points <= 0) return;

            spacing.Line = Invariant((int)Math.Round(points * 20));

            // Travada em "pelo menos", e não em "exatamente": as duas voltam do
            // leitor como pontos, e a que preserva conteúdo é esta — `exact`
            // corta o que não couber na altura declarada.
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

    private void ApplyShading(ParagraphProperties properties, Node node)
    {
        if (Attr.String(node, "background") is not { } background) return;

        if (ColorValue.Hex(background) is not { } fill)
        {
            inventory.NoteLoss($"cor de fundo \"{background}\"");
            return;
        }

        properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
    }

    /// <remarks>
    /// Ausente quer dizer desligado, e desligar é apagar o elemento — mas só
    /// quando ele estava ligado. Um `w:keepNext w:val="false"` existe para
    /// **desligar** o que o estilo liga, e removê-lo faria o parágrafo voltar a
    /// grudar no seguinte.
    /// </remarks>
    private static void ApplyKeepNext(ParagraphProperties properties, Node node)
    {
        if (Attr.Bool(node, "keepNext"))
        {
            properties.KeepNext = new KeepNext();
            return;
        }

        if (RunReader.IsOn(properties.KeepNext)) properties.KeepNext = null;
    }

    /// <summary>
    /// A fonte da marca de parágrafo (`w:pPr/w:rPr`).
    /// </summary>
    /// <remarks>
    /// É ela que o Word usa para medir a linha e para dar altura ao parágrafo
    /// vazio, e é de lá que o leitor tira a fonte do bloco. O resto da marca —
    /// negrito, cor, idioma — segue intocado no clone do original.
    /// </remarks>
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

    /// <summary>
    /// A ordem em que o `w:rPr` declara os seus filhos.
    /// </summary>
    /// <remarks>
    /// O OOXML é uma sequência, não um conjunto: `w:sz` depois de `w:u` torna o
    /// documento inválido, e o Word o recusa inteiro. A marca de parágrafo é a
    /// única propriedade que a biblioteca não expõe tipada — nas outras o próprio
    /// SDK põe cada elemento no lugar —, e por isso a ordem dela mora aqui.
    /// </remarks>
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

        // Elemento que não está na lista vai para o fim: é o que a extensão de
        // um fornecedor faz, e ela sempre se declara depois do padrão.
        var next = mark.ChildElements.FirstOrDefault(existing =>
        {
            var other = Array.IndexOf(MarkOrder, existing.LocalName);
            return other < 0 || other > rank;
        });

        if (next is null) mark.AppendChild(child);
        else mark.InsertBefore(child, next);
    }

    /// <summary>
    /// A primeira fonte da pilha do CSS.
    /// </summary>
    /// <remarks>
    /// O que sai do leitor é uma pilha — a fonte pedida e a substituta genérica —
    /// e o `w:rFonts` guarda o nome de uma fonte, não uma pilha.
    /// </remarks>
    internal static string? FirstFont(string? stack)
    {
        if (stack is null) return null;
        var first = stack.Split(',')[0].Trim().Trim('\'', '"');
        return first.Length > 0 ? first : null;
    }

    /// <summary>
    /// O número como o atributo o escreve, sem a vírgula de nenhuma região.
    /// </summary>
    /// <remarks>
    /// Chamava-se `Twips`, e mentia: metade dos usos é em meios-pontos (`w:sz`) e
    /// em 240-avos (`w:line`). O que ela faz é formatar, não converter.
    /// </remarks>
    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// A medida como o atributo a escreve, ou a **ausência** do atributo.
    /// </summary>
    /// <remarks>
    /// O tipo de retorno é o que importa: devolvendo `StringValue?`, o `null`
    /// apaga o atributo. Escrito como ternário sobre `string`, o mesmo `null`
    /// atravessa a conversão implícita e chega ao SDK como um `StringValue` de
    /// valor nulo, que ele grava como `w:ind w:right=""` — fora do esquema.
    /// </remarks>
    private static StringValue? Twips(int? value) =>
        value is null ? null : new StringValue(Invariant(value.Value));
}

using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Librevia.Format.Docx;

/// <summary>
/// A equação do Word (OMML, `m:oMath` e `m:oMathPara`) desenhada como MathML (M11).
/// </summary>
/// <remarks>
/// Só para a **tela**: o arquivo continua guardando o OMML do jeito que veio, e é
/// ele que volta na gravação — ver ParagraphWriter. Por isso a conversão pode ser
/// aproximada sem custar nada ao documento: o que ela não sabe desenhar vai para
/// <see cref="Converted.Lossy"/>, a equação fica travada na tela, e o OMML segue
/// intacto.
///
/// O alvo é o MathML Core do Chromium, e não o MathML 3 inteiro: sem `menclose`,
/// sem `mlabeledtr`, sem `mathvariant` além de `normal`. A caixa vira `mrow` com
/// classe (a borda é do CSS), e o negrito, o script e o duplo vêm dos caracteres
/// alfanuméricos matemáticos do Unicode — é o que o Core recomenda no lugar do
/// `mathvariant`.
/// </remarks>
public static class OmmlMath
{
    public static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Ml = "http://www.w3.org/1998/Math/MathML";

    /// <summary>A classe da caixa (`m:borderBox`): a borda é desenhada pelo CSS.</summary>
    public const string BoxClass = "omml-caixa";

    /// <param name="MathMl">O `math` inteiro, serializado.</param>
    /// <param name="Display">É uma equação de exibição (`m:oMathPara`).</param>
    /// <param name="Jc">O alinhamento da de exibição (`m:oMathParaPr/m:jc`), quando declarado.</param>
    /// <param name="Lossy">As construções que a tela não desenha, pelo nome OMML (`m:foo`).</param>
    public sealed record Converted(string MathMl, bool Display, string? Jc, IReadOnlyList<string> Lossy);

    /// <summary>Converte o OuterXml de um `m:oMath` ou de um `m:oMathPara`.</summary>
    public static Converted Convert(string omml)
    {
        XElement root;
        try
        {
            root = XElement.Parse(omml);
        }
        catch (XmlException)
        {
            return new Converted(Empty(false), false, null, ["m:oMath"]);
        }

        var converter = new Converter();
        XElement body;
        var display = root.Name == M + "oMathPara";
        string? jc = null;

        if (display)
        {
            jc = Val(root.Element(M + "oMathParaPr"), "jc");
            foreach (var other in root.Elements().Where(child => child.Name != M + "oMath" && child.Name != M + "oMathParaPr"))
            {
                converter.Lose(other);
            }

            // A de várias linhas (Shift+Enter no Word) é um `m:oMath` por linha:
            // vira uma tabela de uma coluna, alinhada como o parágrafo pede.
            var lines = root.Elements(M + "oMath").ToList();
            body = lines.Count == 1
                ? converter.Row(lines[0])
                : new XElement(
                    Ml + "mtable",
                    new XAttribute("columnalign", jc is "left" or "right" ? jc : "center"),
                    lines.Select(line => new XElement(Ml + "mtr", new XElement(Ml + "mtd", converter.Row(line)))));
        }
        else
        {
            body = converter.Row(root);
        }

        var math = new XElement(Ml + "math", new XAttribute("display", display ? "block" : "inline"), body);
        return new Converted(math.ToString(SaveOptions.DisableFormatting), display, jc, [.. converter.Lossy]);
    }

    private static string Empty(bool display) =>
        new XElement(Ml + "math", new XAttribute("display", display ? "block" : "inline")).ToString(SaveOptions.DisableFormatting);

    /// <summary>O `m:val` do filho `name` de uma propriedade — e nulo quando ele não está lá.</summary>
    private static string? Val(XElement? properties, string name) =>
        properties?.Element(M + name)?.Attribute(M + "val")?.Value;

    /// <summary>O liga-desliga do OMML: presente sem valor é ligado.</summary>
    private static bool On(XElement? properties, string name, bool absent = false)
    {
        var element = properties?.Element(M + name);
        if (element is null) return absent;
        return element.Attribute(M + "val")?.Value is not ("0" or "off" or "false");
    }

    private sealed class Converter
    {
        public SortedSet<string> Lossy { get; } = new(StringComparer.Ordinal);

        public void Lose(XElement element) =>
            Lossy.Add((element.Name.Namespace == M ? "m:" : element.Name.Namespace == W ? "w:" : string.Empty) +
                      element.Name.LocalName);

        /// <summary>O conteúdo de um argumento (`m:e`, `m:num`…) num `mrow` — os scripts pedem um filho só.</summary>
        public XElement Row(XElement? container) => new(Ml + "mrow", container is null ? [] : Children(container));

        private IEnumerable<XElement> Children(XElement container)
        {
            foreach (var child in container.Elements())
            {
                if (child.Name.Namespace == W)
                {
                    switch (child.Name.LocalName)
                    {
                        // A revisão dentro da equação: o inserido aparece, o excluído não.
                        case "ins":
                            foreach (var inner in Children(child)) yield return inner;
                            break;
                        case "del" or "rPr" or "bookmarkStart" or "bookmarkEnd" or "proofErr" or
                            "commentRangeStart" or "commentRangeEnd" or "permStart" or "permEnd":
                            break;
                        default:
                            Lose(child);
                            break;
                    }

                    continue;
                }

                if (child.Name.Namespace != M)
                {
                    Lose(child);
                    continue;
                }

                // As propriedades (`m:fPr`, `m:ctrlPr`…) são lidas por quem as tem.
                if (child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)) continue;

                foreach (var element in Element(child)) yield return element;
            }
        }

        private IEnumerable<XElement> Element(XElement element)
        {
            var properties = element.Element(M + element.Name.LocalName + "Pr");
            XElement Arg(string name) => Row(element.Element(M + name));

            switch (element.Name.LocalName)
            {
                case "r":
                    foreach (var token in Run(element)) yield return token;
                    break;

                case "f":
                    switch (Val(properties, "type"))
                    {
                        case "lin" or "skw":
                            yield return new XElement(Ml + "mrow", Arg("num"), Mo("/"), Arg("den"));
                            break;
                        case "noBar":
                            yield return new XElement(Ml + "mfrac", new XAttribute("linethickness", "0"), Arg("num"), Arg("den"));
                            break;
                        default:
                            yield return new XElement(Ml + "mfrac", Arg("num"), Arg("den"));
                            break;
                    }

                    break;

                case "rad":
                    var degree = element.Element(M + "deg");
                    yield return On(properties, "degHide") || degree is null || !degree.Elements().Any(IsContent)
                        ? new XElement(Ml + "msqrt", Arg("e"))
                        : new XElement(Ml + "mroot", Arg("e"), Arg("deg"));
                    break;

                case "sSub":
                    yield return new XElement(Ml + "msub", Arg("e"), Arg("sub"));
                    break;
                case "sSup":
                    yield return new XElement(Ml + "msup", Arg("e"), Arg("sup"));
                    break;
                case "sSubSup":
                    yield return new XElement(Ml + "msubsup", Arg("e"), Arg("sub"), Arg("sup"));
                    break;
                case "sPre":
                    yield return new XElement(Ml + "mmultiscripts", Arg("e"), new XElement(Ml + "mprescripts"), Arg("sub"), Arg("sup"));
                    break;

                case "nary":
                    yield return Nary(element, properties);
                    break;

                case "d":
                    yield return Delimited(element, properties);
                    break;

                case "m":
                    yield return new XElement(
                        Ml + "mtable",
                        element.Elements(M + "mr").Select(row => new XElement(
                            Ml + "mtr",
                            row.Elements(M + "e").Select(cell => new XElement(Ml + "mtd", Row(cell))))));
                    break;

                case "eqArr":
                    yield return new XElement(
                        Ml + "mtable",
                        element.Elements(M + "e").Select(row => new XElement(Ml + "mtr", new XElement(Ml + "mtd", Row(row)))));
                    break;

                case "acc":
                    yield return new XElement(
                        Ml + "mover",
                        new XAttribute("accent", "true"),
                        Arg("e"),
                        Mo(Spacing(Val(properties, "chr") ?? "̂"), stretchy: false));
                    break;

                case "bar":
                    yield return Val(properties, "pos") == "top"
                        ? new XElement(Ml + "mover", new XAttribute("accent", "true"), Arg("e"), Mo("‾", stretchy: true))
                        : new XElement(Ml + "munder", new XAttribute("accentunder", "true"), Arg("e"), Mo("‾", stretchy: true));
                    break;

                case "groupChr":
                    var group = Mo(Val(properties, "chr") ?? "⏟", stretchy: true);
                    yield return Val(properties, "pos") == "top"
                        ? new XElement(Ml + "mover", Arg("e"), group)
                        : new XElement(Ml + "munder", Arg("e"), group);
                    break;

                case "limLow":
                    yield return new XElement(Ml + "munder", Arg("e"), Arg("lim"));
                    break;
                case "limUpp":
                    yield return new XElement(Ml + "mover", Arg("e"), Arg("lim"));
                    break;

                // A função (`sin`, `log`…): o nome, o "aplicação de função" invisível
                // e o argumento.
                case "func":
                    yield return new XElement(Ml + "mrow", Arg("fName"), Mo("⁡"), Arg("e"));
                    break;

                case "box":
                    yield return Arg("e");
                    break;

                case "borderBox":
                    // Só a caixa inteira: lados escondidos e riscos o CSS de um `mrow`
                    // não diz, e a equação fica travada.
                    if (properties is not null && properties.Elements().Any(child =>
                            child.Name.LocalName is "hideTop" or "hideBot" or "hideLeft" or "hideRight" or
                                "strikeH" or "strikeV" or "strikeBLTR" or "strikeTLBR" && On(properties, child.Name.LocalName)))
                    {
                        Lose(element);
                    }

                    yield return new XElement(Ml + "mrow", new XAttribute("class", BoxClass), Arg("e"));
                    break;

                case "phant":
                    XElement phantom = Arg("e");
                    var zeroWidth = On(properties, "zeroWid");
                    var zeroAscent = On(properties, "zeroAsc");
                    var zeroDescent = On(properties, "zeroDesc");
                    if (zeroWidth || zeroAscent || zeroDescent)
                    {
                        phantom = new XElement(
                            Ml + "mpadded",
                            zeroWidth ? new XAttribute("width", "0") : null,
                            zeroAscent ? new XAttribute("height", "0") : null,
                            zeroDescent ? new XAttribute("depth", "0") : null,
                            phantom);
                    }

                    yield return On(properties, "show", absent: true) ? phantom : new XElement(Ml + "mphantom", phantom);
                    break;

                default:
                    // O que não sabemos desenhar: o conteúdo dos argumentos, para a
                    // equação não sumir da tela, e a construção na lista.
                    Lose(element);
                    foreach (var inner in Children(element)) yield return inner;
                    break;
            }
        }

        private static bool IsContent(XElement element) =>
            !element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal);

        private XElement Nary(XElement element, XElement? properties)
        {
            var chr = Val(properties, "chr") ?? "∫";
            var integral = chr is "∫" or "∬" or "∭" or "∮" or "∯" or "∰";
            var underOver = (Val(properties, "limLoc") ?? (integral ? "subSup" : "undOvr")) == "undOvr";
            var hideSub = On(properties, "subHide");
            var hideSup = On(properties, "supHide");

            var op = new XElement(Ml + "mo", new XAttribute("largeop", "true"), chr);
            if (!underOver) op.Add(new XAttribute("movablelimits", "false"));

            var sub = Row(element.Element(M + "sub"));
            var sup = Row(element.Element(M + "sup"));
            XElement limited = (hideSub, hideSup) switch
            {
                (true, true) => op,
                (true, false) => new XElement(Ml + (underOver ? "mover" : "msup"), op, sup),
                (false, true) => new XElement(Ml + (underOver ? "munder" : "msub"), op, sub),
                _ => new XElement(Ml + (underOver ? "munderover" : "msubsup"), op, sub, sup),
            };

            return new XElement(Ml + "mrow", limited, Row(element.Element(M + "e")));
        }

        private XElement Delimited(XElement element, XElement? properties)
        {
            // Ausente é o padrão; o valor vazio é "sem delimitador" — é como o Word
            // grava o colchete de um lado só.
            var open = properties?.Element(M + "begChr") is { } begin ? begin.Attribute(M + "val")?.Value ?? string.Empty : "(";
            var close = properties?.Element(M + "endChr") is { } end ? end.Attribute(M + "val")?.Value ?? string.Empty : ")";
            var separator = properties?.Element(M + "sepChr") is { } sep ? sep.Attribute(M + "val")?.Value ?? string.Empty : "|";

            var row = new XElement(Ml + "mrow");
            if (open.Length > 0) row.Add(Fence(open));
            var first = true;
            foreach (var argument in element.Elements(M + "e"))
            {
                if (!first && separator.Length > 0) row.Add(new XElement(Ml + "mo", new XAttribute("separator", "true"), separator));
                first = false;
                row.Add(Row(argument));
            }

            if (close.Length > 0) row.Add(Fence(close));
            return row;
        }

        private static XElement Fence(string chr) =>
            new(Ml + "mo", new XAttribute("fence", "true"), new XAttribute("stretchy", "true"), chr);

        private static XElement Mo(string text, bool? stretchy = null) =>
            new(Ml + "mo", stretchy is null ? null : new XAttribute("stretchy", stretchy.Value ? "true" : "false"), text);

        /// <summary>O acento combinante do OMML no caractere que fica sozinho em cima da base.</summary>
        private static string Spacing(string chr) => chr switch
        {
            "̀" => "`",
            "́" => "´",
            "̂" => "^",
            "̃" => "~",
            "̄" or "̅" => "¯",
            "̆" => "˘",
            "̇" => "˙",
            "̈" => "¨",
            "̌" => "ˇ",
            "⃖" => "←",
            "⃗" => "→",
            "⃡" => "↔",
            _ => chr,
        };

        // --- o texto ---------------------------------------------------------

        /// <summary>
        /// O texto de um `m:r` em fichas: número (`mn`), identificador (`mi`) e
        /// operador (`mo`); o texto normal (`m:nor`) inteiro num `mtext`.
        /// </summary>
        private static IEnumerable<XElement> Run(XElement run)
        {
            var text = string.Concat(run.Elements(M + "t").Select(t => t.Value));
            if (text.Length == 0) yield break;

            var properties = run.Element(M + "rPr");
            if (On(properties, "nor"))
            {
                yield return new XElement(Ml + "mtext", text);
                yield break;
            }

            var style = Val(properties, "sty");
            var script = Val(properties, "scr") ?? "roman";
            // Itálico é o padrão das letras no Word; `p` é reto.
            var upright = style is "p" or "b";
            var plain = script == "roman" && style is null or "i" or "p";

            var runes = text.EnumerateRunes().ToList();
            for (var i = 0; i < runes.Count;)
            {
                var rune = runes[i];
                if (Rune.IsWhiteSpace(rune))
                {
                    i++;
                    continue;
                }

                if (Rune.IsDigit(rune))
                {
                    var number = new StringBuilder();
                    while (i < runes.Count &&
                           (Rune.IsDigit(runes[i]) ||
                            (runes[i].Value is '.' or ',' && i + 1 < runes.Count && Rune.IsDigit(runes[i + 1]))))
                    {
                        number.Append(Styled(runes[i], style, script));
                        i++;
                    }

                    yield return new XElement(Ml + "mn", number.ToString());
                    continue;
                }

                if (Rune.IsLetter(rune) && rune.Value != '∞')
                {
                    // Reto (o nome de função, o texto `p`) ou de outro alfabeto: as
                    // letras seguidas são um identificador só. Itálico simples: uma
                    // letra por `mi`, que o MathML inclina sozinho.
                    if (plain && !upright)
                    {
                        yield return new XElement(Ml + "mi", rune.ToString());
                        i++;
                        continue;
                    }

                    var word = new StringBuilder();
                    while (i < runes.Count && Rune.IsLetter(runes[i]))
                    {
                        word.Append(Styled(runes[i], style, script));
                        i++;
                    }

                    var identifier = new XElement(Ml + "mi", word.ToString());
                    if (plain) identifier.Add(new XAttribute("mathvariant", "normal"));
                    yield return identifier;
                    continue;
                }

                if (rune.Value == '∞')
                {
                    yield return new XElement(Ml + "mi", rune.ToString());
                    i++;
                    continue;
                }

                yield return new XElement(Ml + "mo", rune.ToString());
                i++;
            }
        }

        /// <summary>
        /// A letra ou o algarismo no alfabeto matemático do Unicode (U+1D400…), que
        /// é como o MathML Core faz negrito, script, fraktur e duplo.
        /// </summary>
        private static string Styled(Rune rune, string? style, string script)
        {
            var c = rune.Value;
            var upper = c is >= 'A' and <= 'Z';
            var lower = c is >= 'a' and <= 'z';
            var digit = c is >= '0' and <= '9';
            if (!upper && !lower && !digit) return rune.ToString();

            var bold = style is "b" or "bi";
            var italic = style is null or "i" or "bi";

            int? letters = script switch
            {
                "script" => bold ? 0x1D4D0 : 0x1D49C,
                "fraktur" => bold ? 0x1D56C : 0x1D504,
                "double-struck" => 0x1D538,
                "sans-serif" => (bold, italic) switch
                {
                    (true, true) => 0x1D63C,
                    (true, false) => 0x1D5D4,
                    (false, true) => 0x1D608,
                    _ => 0x1D5A0,
                },
                "monospace" => 0x1D670,
                // Romano: o itálico simples e o reto ficam como estão — o `mi`
                // inclina a letra sozinho, e `mathvariant="normal"` a endireita.
                _ => (bold, italic) switch
                {
                    (true, true) => 0x1D468,
                    (true, false) => 0x1D400,
                    _ => null,
                },
            };

            int? digits = script switch
            {
                "double-struck" => 0x1D7D8,
                "sans-serif" => bold ? 0x1D7EC : 0x1D7E2,
                "monospace" => 0x1D7F6,
                "roman" when bold => 0x1D7CE,
                _ => null,
            };

            if (digit) return digits is { } zero ? char.ConvertFromUtf32(zero + (c - '0')) : rune.ToString();
            if (letters is not { } start) return rune.ToString();

            if (Hole(script, bold, c) is { } hole) return char.ConvertFromUtf32(hole);
            return char.ConvertFromUtf32(start + (upper ? c - 'A' : 26 + c - 'a'));
        }

        /// <summary>As letras que o Unicode já tinha no bloco Letterlike e deixou de fora do alfabeto.</summary>
        private static int? Hole(string script, bool bold, int c) => (script, bold) switch
        {
            ("script", false) => c switch
            {
                'B' => 0x212C, 'E' => 0x2130, 'F' => 0x2131, 'H' => 0x210B, 'I' => 0x2110, 'L' => 0x2112,
                'M' => 0x2133, 'R' => 0x211B, 'e' => 0x212F, 'g' => 0x210A, 'o' => 0x2134, _ => null,
            },
            ("fraktur", false) => c switch
            {
                'C' => 0x212D, 'H' => 0x210C, 'I' => 0x2111, 'R' => 0x211C, 'Z' => 0x2128, _ => null,
            },
            ("double-struck", _) => c switch
            {
                'C' => 0x2102, 'H' => 0x210D, 'N' => 0x2115, 'P' => 0x2119, 'Q' => 0x211A, 'R' => 0x211D,
                'Z' => 0x2124, _ => null,
            },
            _ => null,
        };
    }

}

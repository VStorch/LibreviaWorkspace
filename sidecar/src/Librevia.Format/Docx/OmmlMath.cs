using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Librevia.Format.Docx;

/// <summary>
/// OMML → MathML, for the **screen** only: the OMML goes back to the file as it came
/// (ParagraphWriter), and what the conversion does not draw goes to <see cref="Converted.Lossy"/>,
/// locking the equation. The target is Chromium's MathML Core: a box becomes an <c>mrow</c> with a
/// class, and bold, script and double-struck come from Unicode's mathematical alphabets instead of
/// <c>mathvariant</c>.
/// </summary>
public static partial class OmmlMath
{
    public static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Ml = "http://www.w3.org/1998/Math/MathML";

    /// <summary>CSS draws the <c>m:borderBox</c> border.</summary>
    public const string BoxClass = "omml-caixa";

    /// <param name="MathMl">The whole `math`, serialized.</param>
    /// <param name="Display">It is a display equation (`m:oMathPara`).</param>
    /// <param name="Jc">The display equation's alignment (`m:oMathParaPr/m:jc`), when
    /// declared.</param>
    /// <param name="Lossy">The constructs the screen does not draw, by OMML name (`m:foo`).</param>
    public sealed record Converted(string MathMl, bool Display, string? Jc, IReadOnlyList<string> Lossy);

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

            // A multi-line one (Shift+Enter in Word) becomes a one-column table, with the paragraph
            // alignment.
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

    /// <summary>Null when the child is not there.</summary>
    private static string? Val(XElement? properties, string name) =>
        properties?.Element(M + name)?.Attribute(M + "val")?.Value;

    /// <summary>Present without a value means on.</summary>
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

        /// <summary>In an <c>mrow</c>: scripts ask for a single child.</summary>
        public XElement Row(XElement? container) => new(Ml + "mrow", container is null ? [] : Children(container));

        private IEnumerable<XElement> Children(XElement container)
        {
            foreach (var child in container.Elements())
            {
                if (child.Name.Namespace == W)
                {
                    switch (child.Name.LocalName)
                    {
                        // Insertions show, deletions do not.
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

                // The name, the invisible "function application" and the argument.
                case "func":
                    yield return new XElement(Ml + "mrow", Arg("fName"), Mo("⁡"), Arg("e"));
                    break;

                case "box":
                    yield return Arg("e");
                    break;

                case "borderBox":
                    // Hidden sides and strikes are something a CSS `mrow` cannot say: the equation
                    // locks.
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

                // A loose argument only appears inside an unknown construct: the content goes on,
                // and the list keeps the construct.
                case "e" or "num" or "den" or "sub" or "sup" or "deg" or "lim" or "fName" or "mr":
                    foreach (var inner in Children(element)) yield return inner;
                    break;

                default:
                    // The content goes on, so the equation does not vanish, and the construct goes
                    // to the list.
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
            // Absent is the default; empty means "no delimiter", Word's one-sided bracket.
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


        /// <summary>
        /// Number (<c>mn</c>), identifier (<c>mi</c>) and operator (<c>mo</c>); <c>m:nor</c> in an
        /// <c>mtext</c>.
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
            // Italic is Word's default for letters; `p` is upright.
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
                    // Upright or from another alphabet, consecutive letters are one identifier;
                    // plain italic, one letter per `mi`, which MathML slants by itself.
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
        /// Unicode's mathematical alphabet (U+1D400…) is how MathML Core does bold, script and
        /// double-struck.
        /// </summary>
        internal static string Styled(Rune rune, string? style, string script)
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
                // Roman: an `mi` slants by itself, and `mathvariant="normal"` straightens it.
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

        /// <summary>Letters Unicode already had in the Letterlike block.</summary>
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

/// <summary>
/// MathML → OMML, only for an equation without <c>omml</c> (new or edited). Accepts the MathML of
/// <see cref="Convert"/> and Temml's, which differ in detail (the sum body, the function wrapper);
/// what it does not recognize goes on as content.
/// </summary>
public static partial class OmmlMath
{
    /// <summary>As Word writes it.</summary>
    private const string MathFont = "Cambria Math";

    /// <summary>Null when the text is not a well-formed <c>math</c>.</summary>
    public static string? ToOmml(string mathMl, bool display, string? jc)
    {
        XElement root;
        try
        {
            root = XElement.Parse(mathMl);
        }
        catch (XmlException)
        {
            return null;
        }

        if (root.Name.LocalName != "math") return null;

        var writer = new Writer();
        var children = root.Elements().Where(child => child.Name.LocalName != "annotation").ToList();

        // A one-column table is a multi-line `m:oMathPara` (see Convert).
        List<XElement> lines;
        if (display && children.Count == 1 && Unwrapped(children[0]) is { } table &&
            table.Name.LocalName == "mtable" && table.Attribute("columnalign") is not null &&
            table.Elements().All(row => row.Elements().Count() == 1))
        {
            lines = [.. table.Elements().Select(row => new XElement(M + "oMath", writer.Row(row.Elements().First())))];
        }
        else
        {
            lines = [new XElement(M + "oMath", writer.Row(root))];
        }

        XElement result;
        if (display)
        {
            var align = jc is "left" or "right" or "center" or "centerGroup" ? jc : "center";
            result = new XElement(
                M + "oMathPara",
                new XElement(M + "oMathParaPr", new XElement(M + "jc", new XAttribute(M + "val", align))),
                lines);
        }
        else
        {
            result = lines[0];
        }

        result.Add(new XAttribute(XNamespace.Xmlns + "m", M.NamespaceName));
        result.Add(new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName));
        return result.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement Unwrapped(XElement element)
    {
        while (element.Name.LocalName == "mrow" && element.Attribute("class") is null && element.Elements().Count() == 1)
        {
            element = element.Elements().First();
        }

        return element;
    }

    private const string NaryChars = "∑∏∐∫∬∭∮∯∰⋀⋁⋂⋃⨀⨁⨂⨄⨆";

    /// <summary>
    /// Temml's n-ary body comes as siblings, up to the first relation or sum operator.
    /// </summary>
    private const string BodyStops = "=+-−±∓<>≤≥≠≈≡∼≃≅∝→←↔⇒⇐⇔∈∉⊂⊃⊆⊇,;";

    private static readonly Dictionary<string, string> Accents = new(StringComparer.Ordinal)
    {
        ["^"] = "̂", ["ˆ"] = "̂", ["̂"] = "̂",
        ["¯"] = "̅", ["‾"] = "̅", ["̄"] = "̄", ["̅"] = "̅",
        ["→"] = "⃗", ["⃗"] = "⃗", ["←"] = "⃖", ["⃖"] = "⃖", ["↔"] = "⃡", ["⃡"] = "⃡",
        ["˙"] = "̇", ["̇"] = "̇", ["¨"] = "̈", ["̈"] = "̈",
        ["~"] = "̃", ["˜"] = "̃", ["̃"] = "̃",
        ["ˇ"] = "̌", ["̌"] = "̌", ["´"] = "́", ["́"] = "́", ["`"] = "̀", ["̀"] = "̀",
        ["˘"] = "̆", ["̆"] = "̆",
    };

    private static readonly Dictionary<int, (char Plain, string? Script, string? Style)> Alphabets = BuildAlphabets();

    private static Dictionary<int, (char, string?, string?)> BuildAlphabets()
    {
        var map = new Dictionary<int, (char, string?, string?)>();
        string[] scripts = ["roman", "script", "fraktur", "double-struck", "sans-serif", "monospace"];
        // The simplest style first: an alphabet without bold gets no `m:sty`.
        string?[] styles = [null, "p", "b", "bi"];
        foreach (var script in scripts)
        {
            foreach (var style in styles)
            {
                foreach (var c in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789")
                {
                    var styled = Converter.Styled(new Rune(c), style, script);
                    if (styled.Length == 1 && styled[0] == c) continue;
                    var code = char.ConvertToUtf32(styled, 0);
                    map.TryAdd(code, (c, script == "roman" ? null : script, style));
                }
            }
        }

        // Roman italic, which only Temml's `\mathit` uses.
        for (var c = 'A'; c <= 'Z'; c++) map.TryAdd(0x1D434 + (c - 'A'), (c, null, "i"));
        for (var c = 'a'; c <= 'z'; c++) map.TryAdd(0x1D44E + (c - 'a'), (c, null, "i"));
        map.TryAdd(0x210E, ('h', null, "i"));
        return map;
    }

    private sealed class Writer
    {
        public List<XElement> Row(XElement? container)
        {
            var output = new List<XElement>();
            if (container is null) return output;
            if (container.Name.LocalName is "mi" or "mn" or "mo" or "mtext" or "ms")
            {
                AddToken(output, container);
                return output;
            }

            AddItems(output, Flatten(container.Elements()));
            return output;
        }

        private List<XElement> Arg(string name, XElement? content) =>
            [new XElement(M + name, Argument(content))];

        /// <summary>An <c>mrow</c> is the group; anything else, a single item.</summary>
        private List<XElement> Argument(XElement? content)
        {
            var output = new List<XElement>();
            if (content is not null) AddItems(output, Flatten([content]));
            return output;
        }

        /// <summary>
        /// Temml's <c>mrow</c> ending in U+2061 is the function name, and the argument is the
        /// sibling.
        /// </summary>
        private static List<XElement> Flatten(IEnumerable<XElement> children)
        {
            var items = new List<XElement>();
            foreach (var child in children)
            {
                var name = child.Name.LocalName;
                if (name is "mspace" or "none" or "annotation" or "annotation-xml") continue;
                if (name is "mstyle" or "mpadded" or "merror" or "semantics")
                {
                    var inner = name == "semantics" ? child.Elements().Take(1) : child.Elements();
                    // Temml's single-child `mpadded` is `\mathrm{abc}`.
                    items.AddRange(Flatten(inner));
                    continue;
                }

                if (name == "mrow" && child.Attribute("class") is null && !Fenced(child))
                {
                    // A single-child wrapper is not a group.
                    var significant = child.Elements().Where(e => e.Name.LocalName != "mspace").ToList();
                    if (significant.Count == 1 || (significant.Count > 0 && IsApply(significant[^1])))
                    {
                        items.AddRange(Flatten(significant));
                        continue;
                    }
                }

                items.Add(child);
            }

            return items;
        }

        private void AddItems(List<XElement> output, List<XElement> items)
        {
            for (var i = 0; i < items.Count; i++)
            {
                var item = items[i];

                // The n-ary body is the next `mrow` (ours) or the siblings (Temml's).
                if (NaryOf(item) is { } nary)
                {
                    var body = new List<XElement>();
                    if (i + 1 < items.Count && IsPlainRow(items[i + 1]))
                    {
                        body.AddRange(Argument(items[++i]));
                    }
                    else
                    {
                        var taken = new List<XElement>();
                        while (i + 1 < items.Count && !IsBodyStop(items[i + 1]) && NaryOf(items[i + 1]) is null)
                        {
                            taken.Add(items[++i]);
                        }

                        AddItems(body, taken);
                    }

                    output.Add(Nary(nary, body));
                    continue;
                }

                if (i + 1 < items.Count && IsApply(items[i + 1]))
                {
                    var name = Argument(item);
                    var argument = new List<XElement>();
                    var next = i + 2;
                    if (next < items.Count)
                    {
                        argument.AddRange(Argument(items[next]));
                        i = next;
                    }
                    else
                    {
                        i++;
                    }

                    output.Add(new XElement(M + "func", new XElement(M + "fName", name), new XElement(M + "e", argument)));
                    continue;
                }

                if (IsApply(item) || IsInvisible(item)) continue;

                // Temml's `{}_a^b X`: an empty-base script followed by the base.
                if (item.Name.LocalName is "msub" or "msup" or "msubsup" && EmptyBase(item) && i + 1 < items.Count)
                {
                    var scripts = item.Elements().ToList();
                    var (sub, sup) = item.Name.LocalName switch
                    {
                        "msub" => (scripts.ElementAtOrDefault(1), (XElement?)null),
                        "msup" => (null, scripts.ElementAtOrDefault(1)),
                        _ => (scripts.ElementAtOrDefault(1), scripts.ElementAtOrDefault(2)),
                    };
                    output.Add(new XElement(M + "sPre", Arg("sub", sub), Arg("sup", sup), Arg("e", items[++i])));
                    continue;
                }

                if (item.Name.LocalName is "mi" or "mn" or "mo" or "mtext" or "ms")
                {
                    AddToken(output, item);
                    continue;
                }

                output.AddRange(Element(item));
            }
        }

        private IEnumerable<XElement> Element(XElement element)
        {
            var args = element.Elements().ToList();
            XElement? At(int index) => args.ElementAtOrDefault(index);

            switch (element.Name.LocalName)
            {
                case "mrow":
                    if (Fenced(element))
                    {
                        yield return Delimited(element);
                    }
                    else if (element.Attribute("class")?.Value.Split(' ').Contains(BoxClass) == true)
                    {
                        yield return new XElement(M + "borderBox", new XElement(M + "e", Row(element)));
                    }
                    else
                    {
                        foreach (var inner in Row(element)) yield return inner;
                    }

                    break;

                case "mfrac":
                    var bar = element.Attribute("linethickness")?.Value is "0" or "0px" or "0pt" or "0em";
                    yield return new XElement(
                        M + "f",
                        bar ? new XElement(M + "fPr", new XElement(M + "type", new XAttribute(M + "val", "noBar"))) : null,
                        Arg("num", At(0)),
                        Arg("den", At(1)));
                    break;

                case "msqrt":
                    yield return new XElement(
                        M + "rad",
                        new XElement(M + "radPr", new XElement(M + "degHide", new XAttribute(M + "val", "1"))),
                        new XElement(M + "deg"),
                        new XElement(M + "e", Row(element)));
                    break;

                case "mroot":
                    yield return new XElement(M + "rad", Arg("deg", At(1)), Arg("e", At(0)));
                    break;

                case "msub":
                    yield return new XElement(M + "sSub", Arg("e", At(0)), Arg("sub", At(1)));
                    break;
                case "msup":
                    yield return new XElement(M + "sSup", Arg("e", At(0)), Arg("sup", At(1)));
                    break;
                case "msubsup":
                    yield return new XElement(M + "sSubSup", Arg("e", At(0)), Arg("sub", At(1)), Arg("sup", At(2)));
                    break;

                case "mmultiscripts":
                {
                    var split = args.FindIndex(arg => arg.Name.LocalName == "mprescripts");
                    if (split >= 0)
                    {
                        yield return new XElement(M + "sPre", Arg("sub", At(split + 1)), Arg("sup", At(split + 2)), Arg("e", At(0)));
                    }
                    else
                    {
                        yield return new XElement(M + "sSubSup", Arg("e", At(0)), Arg("sub", At(1)), Arg("sup", At(2)));
                    }

                    break;
                }

                case "munder":
                    yield return Under(element, At(0), At(1));
                    break;
                case "mover":
                    yield return Over(element, At(0), At(1));
                    break;
                case "munderover":
                    yield return new XElement(
                        M + "limUpp",
                        new XElement(M + "e", new XElement(M + "limLow", Arg("e", At(0)), Arg("lim", At(1)))),
                        Arg("lim", At(2)));
                    break;

                case "mtable":
                    yield return Table(element);
                    break;

                case "mphantom":
                    yield return new XElement(
                        M + "phant",
                        new XElement(M + "phantPr", new XElement(M + "show", new XAttribute(M + "val", "0"))),
                        new XElement(M + "e", Row(element)));
                    break;

                // Unknown content goes on as content.
                default:
                    foreach (var inner in Row(element)) yield return inner;
                    break;
            }
        }

        private XElement Under(XElement element, XElement? baseElement, XElement? under)
        {
            if (Single(under) is { } mark && mark.Name.LocalName == "mo")
            {
                var chr = mark.Value;
                if (chr is "‾" or "_" or "̲" or "¯" && (Stretchy(mark) || element.Attribute("accentunder")?.Value == "true"))
                {
                    return new XElement(
                        M + "bar",
                        new XElement(M + "barPr", new XElement(M + "pos", new XAttribute(M + "val", "bot"))),
                        Arg("e", baseElement));
                }

                if (Stretchy(mark)) return GroupChr(chr, "bot", baseElement);
            }

            return new XElement(M + "limLow", Arg("e", baseElement), Arg("lim", under));
        }

        private XElement Over(XElement element, XElement? baseElement, XElement? over)
        {
            if (Single(over) is { } mark && mark.Name.LocalName == "mo")
            {
                var chr = mark.Value;
                if (chr is "‾" or "¯" && Stretchy(mark))
                {
                    return new XElement(
                        M + "bar",
                        new XElement(M + "barPr", new XElement(M + "pos", new XAttribute(M + "val", "top"))),
                        Arg("e", baseElement));
                }

                if (!Stretchy(mark) && Accents.TryGetValue(chr, out var combining))
                {
                    return new XElement(
                        M + "acc",
                        new XElement(M + "accPr", new XElement(M + "chr", new XAttribute(M + "val", combining))),
                        Arg("e", baseElement));
                }

                if (Stretchy(mark)) return GroupChr(chr, "top", baseElement);
            }

            return new XElement(M + "limUpp", Arg("e", baseElement), Arg("lim", over));
        }

        private XElement GroupChr(string chr, string pos, XElement? baseElement) =>
            new(
                M + "groupChr",
                new XElement(
                    M + "groupChrPr",
                    new XElement(M + "chr", new XAttribute(M + "val", chr)),
                    new XElement(M + "pos", new XAttribute(M + "val", pos)),
                    new XElement(M + "vertJc", new XAttribute(M + "val", pos == "top" ? "bot" : "top"))),
                Arg("e", baseElement));

        private XElement Table(XElement table)
        {
            var rows = table.Elements().Where(row => row.Name.LocalName is "mtr" or "mlabeledtr").ToList();
            if (rows.Count > 0 && rows.All(row => row.Elements().Count() <= 1))
            {
                return new XElement(M + "eqArr", rows.Select(row => new XElement(M + "e", Row(row.Elements().FirstOrDefault()))));
            }

            var columns = Math.Max(1, rows.Count == 0 ? 1 : rows.Max(row => row.Elements().Count()));
            return new XElement(
                M + "m",
                new XElement(
                    M + "mPr",
                    new XElement(
                        M + "mcs",
                        new XElement(
                            M + "mc",
                            new XElement(
                                M + "mcPr",
                                new XElement(M + "count", new XAttribute(M + "val", columns)),
                                new XElement(M + "mcJc", new XAttribute(M + "val", "center")))))),
                rows.Select(row =>
                {
                    var cells = row.Elements().ToList();
                    return new XElement(
                        M + "mr",
                        Enumerable.Range(0, columns).Select(index => new XElement(M + "e", Row(cells.ElementAtOrDefault(index)))));
                }));
        }

        private static bool Fenced(XElement row)
        {
            if (row.Name.LocalName != "mrow") return false;
            var children = row.Elements().ToList();
            if (children.Count < 2) return false;
            return IsFence(children[0]) || IsFence(children[^1]);
        }

        private static bool IsFence(XElement element) =>
            element.Name.LocalName == "mo" && element.Attribute("fence")?.Value == "true";

        private XElement Delimited(XElement row)
        {
            var children = row.Elements().ToList();
            var open = IsFence(children[0]) ? children[0].Value : string.Empty;
            var close = children.Count > 1 && IsFence(children[^1]) ? children[^1].Value : string.Empty;
            var inner = children
                .Skip(IsFence(children[0]) ? 1 : 0)
                .Take(children.Count - (IsFence(children[0]) ? 1 : 0) - (IsFence(children[^1]) ? 1 : 0))
                .ToList();

            // Our `separator` and Temml's `\middle`, a stretchy `mo` that is not a fence.
            var arguments = new List<List<XElement>> { new() };
            string? separator = null;
            foreach (var child in inner)
            {
                if (child.Name.LocalName == "mo" &&
                    (child.Attribute("separator")?.Value == "true" || (Stretchy(child) && !IsFence(child))) &&
                    (separator is null || separator == child.Value))
                {
                    separator = child.Value;
                    arguments.Add([]);
                    continue;
                }

                arguments[^1].Add(child);
            }

            var properties = new XElement(M + "dPr", new XElement(M + "begChr", new XAttribute(M + "val", open)));
            if (arguments.Count > 1) properties.Add(new XElement(M + "sepChr", new XAttribute(M + "val", separator ?? "|")));
            properties.Add(new XElement(M + "endChr", new XAttribute(M + "val", close)));

            return new XElement(
                M + "d",
                properties,
                arguments.Select(argument =>
                {
                    var content = new List<XElement>();
                    AddItems(content, Flatten(argument));
                    return new XElement(M + "e", content);
                }));
        }

        private static XElement? NaryOf(XElement element)
        {
            var name = element.Name.LocalName;
            if (name == "mo") return IsNaryOperator(element) ? element : null;
            if (name is "msub" or "msup" or "msubsup" or "munder" or "mover" or "munderover" &&
                element.Elements().FirstOrDefault() is { } first && Unwrapped(first) is var op &&
                op.Name.LocalName == "mo" && IsNaryOperator(op))
            {
                return element;
            }

            return null;
        }

        private static bool IsNaryOperator(XElement mo) =>
            mo.Value.Trim() is { Length: > 0 } text && (NaryChars.Contains(text, StringComparison.Ordinal) && text.Length == 1 ||
                                                       mo.Attribute("largeop")?.Value == "true");

        private XElement Nary(XElement nary, List<XElement> body)
        {
            var args = nary.Elements().ToList();
            var op = nary.Name.LocalName == "mo" ? nary : Unwrapped(args[0]);
            var (limLoc, sub, sup) = nary.Name.LocalName switch
            {
                "msub" => ("subSup", args.ElementAtOrDefault(1), (XElement?)null),
                "msup" => ("subSup", null, args.ElementAtOrDefault(1)),
                "msubsup" => ("subSup", args.ElementAtOrDefault(1), args.ElementAtOrDefault(2)),
                "munder" => ("undOvr", args.ElementAtOrDefault(1), null),
                "mover" => ("undOvr", null, args.ElementAtOrDefault(1)),
                "munderover" => ("undOvr", args.ElementAtOrDefault(1), args.ElementAtOrDefault(2)),
                _ => (op.Attribute("movablelimits")?.Value == "false" || "∫∬∭∮∯∰".Contains(op.Value.Trim(), StringComparison.Ordinal)
                    ? "subSup"
                    : "undOvr", null, null),
            };

            var properties = new XElement(
                M + "naryPr",
                new XElement(M + "chr", new XAttribute(M + "val", op.Value.Trim())),
                new XElement(M + "limLoc", new XAttribute(M + "val", limLoc)));
            if (sub is null) properties.Add(new XElement(M + "subHide", new XAttribute(M + "val", "1")));
            if (sup is null) properties.Add(new XElement(M + "supHide", new XAttribute(M + "val", "1")));

            return new XElement(M + "nary", properties, Arg("sub", sub), Arg("sup", sup), new XElement(M + "e", body));
        }

        private static bool IsPlainRow(XElement element) =>
            element.Name.LocalName == "mrow" && element.Attribute("class") is null && !Fenced(element);

        private static bool IsBodyStop(XElement element) =>
            element.Name.LocalName == "mo" && element.Value.Trim() is { Length: > 0 } text &&
            BodyStops.Contains(text, StringComparison.Ordinal) && text.Length == 1;

        private static bool IsApply(XElement element) => element.Name.LocalName == "mo" && element.Value == "⁡";

        private static bool IsInvisible(XElement element) =>
            element.Name.LocalName == "mo" && element.Value is "⁢" or "⁣" or "⁤" or "";

        private static bool EmptyBase(XElement script) =>
            script.Elements().FirstOrDefault() is { } first && first.Name.LocalName == "mrow" && !first.HasElements &&
            first.Value.Length == 0;

        private static XElement? Single(XElement? element) => element is null ? null : Unwrapped(element);

        private static bool Stretchy(XElement mo) => mo.Attribute("stretchy")?.Value == "true";


        /// <summary>
        /// Merged with the previous one when the properties are the same, as Word writes it.
        /// </summary>
        private static void AddToken(List<XElement> output, XElement token)
        {
            var text = token.Value;
            if (text.Length == 0) return;

            string? script = null;
            string? style = null;
            var normal = false;
            var plain = new StringBuilder();

            if (token.Name.LocalName is "mtext" or "ms")
            {
                normal = true;
                plain.Append(text);
            }
            else
            {
                var styledAlphabet = false;
                foreach (var rune in text.EnumerateRunes())
                {
                    if (Alphabets.TryGetValue(rune.Value, out var letter))
                    {
                        plain.Append(letter.Plain);
                        script = letter.Script;
                        style = letter.Style;
                        styledAlphabet = true;
                    }
                    else if (rune.Value is 0x2061 or 0x2062 or 0x2063 or 0x2064)
                    {
                        continue;
                    }
                    else
                    {
                        plain.Append(rune.ToString());
                    }
                }

                if (plain.Length == 0) return;

                if (!styledAlphabet && token.Name.LocalName == "mi")
                {
                    // A multi-letter `mi`, or `mathvariant="normal"`, is Word's `m:sty p`.
                    var upright = token.Attribute("mathvariant")?.Value == "normal" ||
                                  (token.Attribute("mathvariant") is null && plain.ToString().EnumerateRunes().Count() > 1);
                    var variant = token.Attribute("mathvariant")?.Value;
                    style = upright ? "p" : variant switch
                    {
                        "bold" => "b",
                        "bold-italic" => "bi",
                        _ => null,
                    };
                }
                else if (styledAlphabet && style == "i")
                {
                    // Italic is Word's default.
                    style = null;
                }
            }

            var value = plain.ToString();
            if (output.Count > 0 && output[^1] is { } last && last.Name == M + "r" &&
                Same(last, script, style, normal))
            {
                var t = last.Element(M + "t")!;
                t.Value += value;
                Preserve(t);
                return;
            }

            output.Add(Run(value, script, style, normal));
        }

        private static XElement Run(string text, string? script, string? style, bool normal)
        {
            XElement? properties = null;
            if (normal)
            {
                properties = new XElement(M + "rPr", new XElement(M + "nor"));
            }
            else if (script is not null || style is not null)
            {
                properties = new XElement(
                    M + "rPr",
                    script is null ? null : new XElement(M + "scr", new XAttribute(M + "val", script)),
                    style is null ? null : new XElement(M + "sty", new XAttribute(M + "val", style)));
            }

            var t = new XElement(M + "t", text);
            Preserve(t);
            return new XElement(
                M + "r",
                properties,
                new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", MathFont), new XAttribute(W + "hAnsi", MathFont))),
                t);
        }

        private static void Preserve(XElement t)
        {
            var value = t.Value;
            var needs = value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
            t.SetAttributeValue(XNamespace.Xml + "space", needs ? "preserve" : null);
        }

        private static bool Same(XElement run, string? script, string? style, bool normal)
        {
            var properties = run.Element(M + "rPr");
            var runNormal = properties?.Element(M + "nor") is not null;
            var runScript = properties?.Element(M + "scr")?.Attribute(M + "val")?.Value;
            var runStyle = properties?.Element(M + "sty")?.Attribute(M + "val")?.Value;
            return runNormal == normal && runScript == script && runStyle == style;
        }
    }
}

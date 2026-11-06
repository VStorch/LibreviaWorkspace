using System.Xml;
using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// The equation's MathML comes from the model, which anyone can edit by hand: only MathML Core
/// elements and presentation attributes pass, the same lists as <c>mathml.ts</c>. Read without DTD
/// or resolver; a comment, CDATA, processing instruction, unknown entity, foreign namespace or
/// prefix refuses the whole text.
/// </summary>
internal static class MathMlFilter
{
    public static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    private const string BoxClass = "omml-box";
    private const int MaxDepth = 128;
    private const int MaxNodes = 50_000;

    private static readonly HashSet<string> Elements = new(StringComparer.Ordinal)
    {
        "math", "mrow", "mi", "mn", "mo", "mtext", "mspace", "ms", "mfrac", "msqrt", "mroot", "msub",
        "msup", "msubsup", "munder", "mover", "munderover", "mmultiscripts", "mprescripts", "none",
        "mtable", "mtr", "mtd", "mpadded", "mphantom", "mstyle", "merror", "semantics", "annotation",
    };

    /// <summary>Elements that carry text; elsewhere, text between children is only XML white space.</summary>
    private static readonly HashSet<string> Tokens = new(StringComparer.Ordinal)
    {
        "mi", "mn", "mo", "mtext", "ms", "annotation",
    };

    private static readonly HashSet<string> Attributes = new(StringComparer.Ordinal)
    {
        "display", "displaystyle", "scriptlevel", "mathvariant", "linethickness", "largeop",
        "movablelimits", "stretchy", "symmetric", "fence", "separator", "accent", "accentunder",
        "lspace", "rspace", "minsize", "maxsize", "columnalign", "rowalign", "columnspan", "rowspan",
        "width", "height", "depth", "voffset", "dir", "encoding", "class",
    };

    private static readonly HashSet<string> StackedElements = new(StringComparer.Ordinal)
    {
        "mfrac", "munder", "mover", "munderover", "mtable",
    };

    /// <returns>The filtered <c>math</c> element, or <c>null</c> when the text is not acceptable MathML.</returns>
    public static XElement? Sanitize(string? source)
    {
        if (string.IsNullOrWhiteSpace(source)) return null;
        var parsed = Parse(source);
        return parsed is null || parsed.Name != MathMl + "math" ? null : Clean(parsed);
    }

    private static XElement? Parse(string source)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };
        try
        {
            using var reader = XmlReader.Create(new StringReader(source), settings);
            return Read(reader);
        }
        catch (XmlException)
        {
            return null;
        }
    }

    private static XElement? Read(XmlReader reader)
    {
        var stack = new Stack<XElement>();
        XElement? root = null;
        var count = 0;
        while (reader.Read())
        {
            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    if (!Acceptable(reader) || ++count > MaxNodes) return null;
                    var element = new XElement(MathMl + reader.LocalName, AttributesOf(reader));
                    if (stack.TryPeek(out var parent)) parent.Add(element);
                    else if (root is null) root = element;
                    else return null;
                    if (!reader.IsEmptyElement) stack.Push(element);
                    if (stack.Count > MaxDepth) return null;
                    break;
                case XmlNodeType.EndElement:
                    stack.Pop();
                    break;
                case XmlNodeType.Text or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace:
                    if (stack.TryPeek(out var holder)) holder.Add(new XText(reader.Value));
                    break;
                case XmlNodeType.XmlDeclaration:
                    break;
                default:
                    return null;
            }
        }

        return stack.Count == 0 ? root : null;
    }

    private static bool Acceptable(XmlReader reader) =>
        reader.Prefix.Length == 0 && (reader.NamespaceURI.Length == 0 || reader.NamespaceURI == MathMl.NamespaceName);

    private static List<XAttribute> AttributesOf(XmlReader reader)
    {
        var attributes = new List<XAttribute>();
        while (reader.MoveToNextAttribute())
        {
            if (reader.Prefix.Length == 0 && reader.NamespaceURI.Length == 0) attributes.Add(new XAttribute(reader.LocalName, reader.Value));
        }

        reader.MoveToElement();
        return attributes;
    }

    private static XElement? Clean(XElement element)
    {
        if (!Elements.Contains(element.Name.LocalName)) return null;
        var cleaned = new XElement(element.Name);
        foreach (var attribute in element.Attributes())
        {
            var name = attribute.Name.LocalName;
            if (!Attributes.Contains(name)) continue;
            if (name == "class")
            {
                if (attribute.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Contains(BoxClass)) cleaned.Add(new XAttribute(name, BoxClass));
                continue;
            }

            cleaned.Add(new XAttribute(name, attribute.Value));
        }

        var token = Tokens.Contains(element.Name.LocalName);
        foreach (var child in element.Nodes())
        {
            if (child is XText text && token) cleaned.Add(new XText(text.Value));
            else if (child is XElement inner && Clean(inner) is { } kept) cleaned.Add(kept);
        }

        return cleaned;
    }

    /// <summary>Token text in order, without annotations: what the equation says.</summary>
    public static string TextOf(XElement element) => string.Concat(element.Nodes().Select(node => node switch
    {
        XText text => text.Value,
        XElement { Name.LocalName: "annotation" } => string.Empty,
        XElement inner => TextOf(inner),
        _ => string.Empty,
    }));

    /// <summary>How many levels the equation stacks: fractions, limits and table rows.</summary>
    public static int StackedLevels(XElement element)
    {
        var inner = element.Elements().Select(StackedLevels).DefaultIfEmpty(0).Max();
        if (element.Name.LocalName == "mtable") return inner + Math.Max(0, element.Elements().Count() - 1);
        return inner + (StackedElements.Contains(element.Name.LocalName) ? 1 : 0);
    }
}

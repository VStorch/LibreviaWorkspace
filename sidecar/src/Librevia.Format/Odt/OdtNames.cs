using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>The ODF namespaces the reader looks into.</summary>
internal static class Ns
{
    public static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    public static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    public static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    public static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    public static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    public static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    public static readonly XNamespace XLink = "http://www.w3.org/1999/xlink";
    public static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    public static readonly XNamespace Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
    public static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    public static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    public static readonly XNamespace Loext = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
    public static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    public static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";
    public static readonly XNamespace Xml = XNamespace.Xml;

    private static readonly Dictionary<XNamespace, string> Prefixes = new()
    {
        [Office] = "office",
        [Style] = "style",
        [Text] = "text",
        [Table] = "table",
        [Draw] = "draw",
        [Fo] = "fo",
        [Svg] = "svg",
        [Dc] = "dc",
        [Meta] = "meta",
        [Loext] = "loext",
        ["urn:oasis:names:tc:opendocument:xmlns:form:1.0"] = "form",
        ["urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"] = "presentation",
        ["urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"] = "dr3d",
        ["urn:oasis:names:tc:opendocument:xmlns:chart:1.0"] = "chart",
        [MathMl] = "math",
    };

    /// <summary>The tag as the file names it, for the inventory: <c>text:ruby</c>.</summary>
    public static string Tag(XElement element)
    {
        var name = element.Name;
        return Prefixes.TryGetValue(name.Namespace, out var prefix)
            ? $"{prefix}:{name.LocalName}"
            : element.GetPrefixOfNamespace(name.Namespace) is { Length: > 0 } own
                ? $"{own}:{name.LocalName}"
                : name.LocalName;
    }

    public static string? Attr(this XElement element, XName name) => element.Attribute(name)?.Value;
}

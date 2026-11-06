using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// An element written as a piece of an existing part: with the prefixes its root declares, whatever
/// they are (<c>o:p</c> works as well as <c>text:p</c>). A namespace the root lacks is declared on
/// the piece itself, under a prefix the part does not use.
/// </summary>
internal sealed class OdtXmlOut
{
    private readonly Dictionary<XNamespace, string> _prefixes = [];
    private readonly HashSet<string> _taken = new(StringComparer.Ordinal);

    public OdtXmlOut(XElement root)
    {
        foreach (var declaration in root.Attributes().Where(attribute => attribute.IsNamespaceDeclaration))
        {
            var prefix = declaration.Name.Namespace == XNamespace.None ? string.Empty : declaration.Name.LocalName;
            _taken.Add(prefix);
            _prefixes.TryAdd(declaration.Value, prefix);
        }
    }

    public string Write(XElement element) => Write([element]);

    public string Write(IEnumerable<XElement> elements)
    {
        var text = new StringBuilder();
        foreach (var element in elements) WriteElement(text, element, Missing(element), top: true);
        return text.ToString();
    }

    /// <summary>
    /// The namespaces of the piece the root does not declare, with a fresh prefix each. An attribute
    /// cannot take the default namespace, so that one counts as missing for attributes.
    /// </summary>
    private Dictionary<XNamespace, string> Missing(XElement element)
    {
        var missing = new Dictionary<XNamespace, string>();
        foreach (var node in element.DescendantsAndSelf())
        {
            Need(missing, node.Name.Namespace, attribute: false);
            foreach (var attribute in node.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration))
            {
                Need(missing, attribute.Name.Namespace, attribute: true);
            }
        }

        return missing;
    }

    private void Need(Dictionary<XNamespace, string> missing, XNamespace space, bool attribute)
    {
        if (space == XNamespace.None || space == XNamespace.Xml || missing.ContainsKey(space)) return;
        if (_prefixes.TryGetValue(space, out var prefix) && (prefix.Length > 0 || !attribute)) return;

        var number = missing.Count + 1;
        while (_taken.Contains($"ns{number}")) number++;
        missing[space] = $"ns{number}";
    }

    private void WriteElement(StringBuilder text, XElement element, Dictionary<XNamespace, string> declare, bool top)
    {
        var name = Qualified(element.Name, declare);
        text.Append('<').Append(name);
        foreach (var (space, prefix) in declare.Where(_ => top))
        {
            text.Append(" xmlns:").Append(prefix).Append("=\"").Append(Escape(space.NamespaceName, attribute: true)).Append('"');
        }

        foreach (var attribute in element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration))
        {
            text.Append(' ').Append(Qualified(attribute.Name, declare))
                .Append("=\"").Append(Escape(attribute.Value, attribute: true)).Append('"');
        }

        if (!element.Nodes().Any())
        {
            text.Append("/>");
            return;
        }

        text.Append('>');
        foreach (var node in element.Nodes())
        {
            if (node is XElement child) WriteElement(text, child, declare, top: false);
            else if (node is XText plain) text.Append(Escape(plain.Value, attribute: false));
        }

        text.Append("</").Append(name).Append('>');
    }

    private string Qualified(XName name, Dictionary<XNamespace, string> declare)
    {
        if (name.Namespace == XNamespace.None) return name.LocalName;
        if (name.Namespace == XNamespace.Xml) return "xml:" + name.LocalName;
        var prefix = declare.TryGetValue(name.Namespace, out var local) ? local : _prefixes.GetValueOrDefault(name.Namespace);
        return string.IsNullOrEmpty(prefix) ? name.LocalName : $"{prefix}:{name.LocalName}";
    }

    /// <summary>What XML refuses is dropped, as <c>odt-xml.ts</c> does.</summary>
    private static string Escape(string value, bool attribute)
    {
        var text = new StringBuilder(value.Length);
        foreach (var letter in value)
        {
            if (!XmlConvert.IsXmlChar(letter) && !char.IsSurrogate(letter)) continue;
            text.Append(letter switch
            {
                '&' => "&amp;",
                '<' => "&lt;",
                '>' => "&gt;",
                '"' when attribute => "&quot;",
                '\n' when attribute => "&#10;",
                '\r' => "&#13;",
                '\t' when attribute => "&#9;",
                _ => letter.ToString(),
            });
        }

        return text.ToString();
    }
}

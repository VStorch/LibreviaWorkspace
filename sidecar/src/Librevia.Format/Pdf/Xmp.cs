using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Librevia.Format.Pdf;

/// <summary>
/// The XMP packet that Acrobat reads before the information dictionary. Only the fields that mirror
/// it change; the rest of the packet, wrapper included, stays.
/// </summary>
internal static class Xmp
{
    private static readonly XNamespace Rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace PdfNs = "http://ns.adobe.com/pdf/1.3/";
    private static readonly XNamespace XmpNs = "http://ns.adobe.com/xap/1.0/";

    private static readonly Dictionary<XNamespace, string> Prefixes = new()
    {
        [Dc] = "dc", [PdfNs] = "pdf", [XmpNs] = "xmp",
    };

    private enum Shape
    {
        Simple,
        Alternative,
        Sequence,
    }

    /// <summary>In the order of the information dictionary's keys.</summary>
    private static readonly (string Key, XName Name, Shape Shape)[] Fields =
    [
        ("Title", Dc + "title", Shape.Alternative),
        ("Author", Dc + "creator", Shape.Sequence),
        ("Subject", Dc + "description", Shape.Alternative),
        ("Keywords", PdfNs + "Keywords", Shape.Simple),
        ("Creator", XmpNs + "CreatorTool", Shape.Simple),
        ("Producer", PdfNs + "Producer", Shape.Simple),
        ("CreationDate", XmpNs + "CreateDate", Shape.Simple),
        ("ModDate", XmpNs + "ModifyDate", Shape.Simple),
    ];

    public static XDocument? Parse(byte[] data)
    {
        try
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using var reader = XmlReader.Create(new MemoryStream(data), settings);
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            return null;
        }
    }

    /// <summary>By the information dictionary's key; dates as W3CDTF.</summary>
    public static Dictionary<string, string> Read(XDocument packet)
    {
        var values = new Dictionary<string, string>();
        foreach (var (key, name, _) in Fields)
        {
            var value = Find(packet, name) switch
            {
                XAttribute attribute => attribute.Value,
                XElement element => element.Descendants(Rdf + "li").FirstOrDefault()?.Value ?? element.Value,
                _ => null,
            };
            if (!string.IsNullOrWhiteSpace(value)) values[key] = value.Trim();
        }

        return values;
    }

    /// <summary>Null value: leave; empty: remove. Dates come as W3CDTF.</summary>
    public static void Update(XDocument packet, IReadOnlyDictionary<string, string> changes)
    {
        foreach (var (key, name, shape) in Fields)
        {
            if (!changes.TryGetValue(key, out var value)) continue;
            Set(packet, name, shape, value);
        }

        if (changes.TryGetValue("ModDate", out var modified)) Set(packet, XmpNs + "MetadataDate", Shape.Simple, modified);
    }

    public static byte[] Write(XDocument packet)
    {
        var output = new MemoryStream();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = packet.Declaration is null,
        };
        using (var writer = XmlWriter.Create(output, settings))
        {
            packet.Save(writer);
        }

        return output.ToArray();
    }

    private static XObject? Find(XDocument packet, XName name)
    {
        foreach (var description in packet.Descendants(Rdf + "Description"))
        {
            if (description.Attribute(name) is { } attribute) return attribute;
            if (description.Element(name) is { } element) return element;
        }

        return null;
    }

    private static void Set(XDocument packet, XName name, Shape shape, string value)
    {
        var existing = Find(packet, name);
        if (value.Length == 0)
        {
            Remove(existing);
            return;
        }

        if (existing is XAttribute attribute && shape == Shape.Simple)
        {
            attribute.Value = value;
            return;
        }

        var description = existing?.Parent ?? packet.Descendants(Rdf + "Description").FirstOrDefault();
        Remove(existing);
        if (description is null) throw new PdfException("sidecar.error.xmpUnsupported");

        if (description.GetPrefixOfNamespace(name.Namespace) is null)
        {
            description.SetAttributeValue(XNamespace.Xmlns + Prefixes[name.Namespace], name.NamespaceName);
        }

        description.Add(new XElement(name, Content(shape, value)));
    }

    private static void Remove(XObject? node)
    {
        if (node is XAttribute attribute) attribute.Remove();
        else if (node is XElement element) element.Remove();
    }

    private static object Content(Shape shape, string value) => shape switch
    {
        Shape.Alternative => new XElement(
            Rdf + "Alt",
            new XElement(Rdf + "li", new XAttribute(XNamespace.Xml + "lang", "x-default"), value)),
        Shape.Sequence => new XElement(Rdf + "Seq", new XElement(Rdf + "li", value)),
        _ => value,
    };
}

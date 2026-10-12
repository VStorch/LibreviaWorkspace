using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Librevia.Format;

/// <summary>
/// A package part as raw XML, for the parts the SDK would reserialize whole and drop what it does
/// not know.
/// </summary>
internal static class XmlParts
{
    /// <summary>Null when the XML is malformed.</summary>
    public static XDocument? Read(Stream stream)
    {
        try
        {
            using var reader = XmlReader.Create(
                stream,
                new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            return null;
        }
    }

    public static XDocument? Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return Read(stream);
    }

    public static void Write(Stream stream, XDocument xml)
    {
        // XmlWriter would write `utf-8`; Word and LibreOffice write `UTF-8`.
        var standalone = xml.Declaration?.Standalone is { Length: > 0 } value ? value : "yes";
        var declaration = Encoding.UTF8.GetBytes(
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"{standalone}\"?>\r\n");
        stream.Write(declaration, 0, declaration.Length);
        using var writer = XmlWriter.Create(
            stream,
            new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = true });
        xml.Root!.Save(writer);
    }

    public static byte[] Write(XDocument xml)
    {
        using var stream = new MemoryStream();
        Write(stream, xml);
        return stream.ToArray();
    }
}

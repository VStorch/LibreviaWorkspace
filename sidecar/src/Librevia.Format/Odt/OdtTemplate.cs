using System.IO.Compression;
using System.Text;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// The <c>odt.create</c> request: what lives in <c>styles.xml</c> (page and sections, styles, note
/// numbering, the heading levels the body uses) and in <c>meta.xml</c>.
/// </summary>
public sealed record OdtCreateDto(
    [property: JsonPropertyName("page")] PageSetupDto Page,
    [property: JsonPropertyName("sections")] List<PageSetupDto>? Sections = null,
    [property: JsonPropertyName("styles")] StyleSheetDto? Styles = null,
    [property: JsonPropertyName("notes")] NotesDto? Notes = null,
    [property: JsonPropertyName("generator")] string? Generator = null,
    [property: JsonPropertyName("properties")] PropertiesDto? Properties = null,
    [property: JsonPropertyName("headingLevels")] List<int>? HeadingLevels = null);

/// <summary>
/// The package that plays the original for a document without one, as <see cref="DocxTemplate"/>
/// does for DOCX: an empty <c>office:text</c>, so <c>odt.save</c> writes every block as new. The
/// named styles, master pages and headers go here, because the body only points to them.
/// </summary>
public static class OdtTemplate
{
    /// <summary>The earliest zip date, so the same request gives the same bytes.</summary>
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    private const string OdfVersion = "1.3";
    private const string XmlType = "text/xml";
    private const string DefaultGenerator = "Librevia";

    internal static readonly XNamespace Number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

    public static byte[] Create(OdtCreateDto request)
    {
        var pictures = new OdtPictures(null);
        var styles = OdtTemplateStyles.Write(request, pictures);
        var entries = new List<(string Path, byte[] Bytes, bool Stored)>
        {
            (OdtPackage.MimeTypePath, Encoding.ASCII.GetBytes(OdtPackage.TextMimeType), true),
            (OdtPackage.ContentPath, Xml(Content()), false),
            (OdtPackage.StylesPath, Xml(styles), false),
            (OdtPackage.MetaPath, Xml(Meta(request.Generator, request.Properties)), false),
        };
        entries.AddRange(pictures.Added.Select(added => (added.Path, added.Bytes, true)));
        entries.Add((OdtPackage.ManifestPath, Xml(Manifest(pictures)), false));
        return Zip(entries);
    }

    /// <summary>Every namespace a save may write, declared on the root, so pieces carry none.</summary>
    internal static XElement Root(XName name) => new(
        name,
        new XAttribute(XNamespace.Xmlns + "office", Ns.Office.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "style", Ns.Style.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "text", Ns.Text.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "table", Ns.Table.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "draw", Ns.Draw.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "fo", Ns.Fo.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "xlink", Ns.XLink.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "dc", Ns.Dc.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "meta", Ns.Meta.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "number", Number.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "svg", Ns.Svg.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "loext", Ns.Loext.NamespaceName),
        new XAttribute(Ns.Office + "version", OdfVersion));

    private static XDocument Content()
    {
        var root = Root(Ns.Office + "document-content");
        root.Add(
            new XElement(Ns.Office + "automatic-styles"),
            new XElement(Ns.Office + "body", new XElement(Ns.Office + "text")));
        return new XDocument(root);
    }

    /// <summary>
    /// The generator, and what only a new file says: the author of the last change falls back to
    /// the creator, and the editing time. <c>odt.save</c> writes the other properties.
    /// </summary>
    private static XDocument Meta(string? generator, PropertiesDto? properties)
    {
        var meta = new XElement(Ns.Office + "meta", new XElement(Ns.Meta + "generator", generator ?? DefaultGenerator));
        var author = properties?.LastModifiedBy is { Length: > 0 } last ? last : properties?.Creator;
        if (author is { Length: > 0 }) meta.Add(new XElement(Ns.Dc + "creator", author));
        if (properties?.TotalTime is > 0 and var minutes) meta.Add(new XElement(Ns.Meta + "editing-duration", $"PT{minutes}M"));
        var root = Root(Ns.Office + "document-meta");
        root.Add(meta);
        return new XDocument(root);
    }

    private static XDocument Manifest(OdtPictures pictures)
    {
        static XElement Entry(string path, string type) => new(
            Ns.Manifest + "file-entry",
            new XAttribute(Ns.Manifest + "full-path", path),
            new XAttribute(Ns.Manifest + "media-type", type));

        var whole = Entry("/", OdtPackage.TextMimeType);
        whole.Add(new XAttribute(Ns.Manifest + "version", OdfVersion));
        return new XDocument(new XElement(
            Ns.Manifest + "manifest",
            new XAttribute(XNamespace.Xmlns + "manifest", Ns.Manifest.NamespaceName),
            new XAttribute(Ns.Manifest + "version", OdfVersion),
            whole,
            Entry(OdtPackage.ContentPath, XmlType),
            Entry(OdtPackage.StylesPath, XmlType),
            Entry(OdtPackage.MetaPath, XmlType),
            pictures.ManifestEntries()));
    }

    private static byte[] Xml(XDocument document) => XmlParts.Write(document);

    /// <summary><c>mimetype</c> first and stored, so a reader finds the type at a fixed offset.</summary>
    private static byte[] Zip(List<(string Path, byte[] Bytes, bool Stored)> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (path, bytes, stored) in entries)
            {
                var entry = zip.CreateEntry(path, stored ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
                entry.LastWriteTime = ZipEpoch;
                using var target = entry.Open();
                target.Write(bytes);
            }
        }

        return output.ToArray();
    }
}

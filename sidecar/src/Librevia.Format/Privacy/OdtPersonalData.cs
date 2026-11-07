using System.IO.Compression;
using System.Xml.Linq;
using Librevia.Format.Odt;

namespace Librevia.Format.Privacy;

/// <summary>
/// <see cref="PersonalData"/> for an ODT or OTT, with the same report and categories. The package is
/// read with the defences of <see cref="OdtPackage"/>; only the parts that change are written again,
/// <c>mimetype</c> first and stored, the rest in the original order.
/// </summary>
internal sealed class OdtPersonalData
{
    private const string ThumbnailPath = "Thumbnails/thumbnail.png";
    private const string ThumbnailFolder = "Thumbnails/";
    private const string PicturesFolder = "Pictures/";
    private const string XmlType = "text/xml";
    private const int MaxMimeTypeBytes = 256;
    private const string ContentTypesPart = "[Content_Types].xml";

    private static readonly XName[] MetaNames = [Ns.Meta + "initial-creator", Ns.Dc + "creator", Ns.Meta + "printed-by"];
    private static readonly string[] NamedUserDefined = [OdtProperties.Company, OdtProperties.Manager];
    private static readonly XName[] AuthoredElements = [Ns.Office + "annotation", Ns.Office + "change-info"];
    private static readonly string[] NamedParts = [OdtPackage.MetaPath, OdtPackage.ContentPath, OdtPackage.StylesPath];

    private readonly OdtPackage _package;
    private readonly Dictionary<string, XDocument> _xml = new(StringComparer.Ordinal);
    private readonly HashSet<string> _edited = new(StringComparer.Ordinal);
    private readonly Dictionary<string, byte[]> _replaced = new(StringComparer.Ordinal);
    private readonly HashSet<string> _removed = new(StringComparer.Ordinal);

    private OdtPersonalData(OdtPackage package) => _package = package;

    /// <summary>
    /// A <c>mimetype</c> of the ODF text family, so a master document gets the ODT error; without one, a zip with the parts of an ODF package and not the
    /// Open XML content types.
    /// </summary>
    public static bool IsOdt(byte[] bytes)
    {
        try
        {
            using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
            var mimeType = archive.GetEntry(OdtPackage.MimeTypePath);
            if (mimeType is null) return IsUntypedOdt(archive);
            if (mimeType.Length > MaxMimeTypeBytes) return false;

            using var stream = mimeType.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return OdtPackage.MimeTypeIn(copy.ToArray()).StartsWith(OdtPackage.TextMimeType, StringComparison.Ordinal);
        }
        catch (Exception problem) when (problem is InvalidDataException or IOException or NotSupportedException)
        {
            return false;
        }
    }

    private static bool IsUntypedOdt(ZipArchive archive) =>
        archive.GetEntry(ContentTypesPart) is null &&
        (archive.GetEntry(OdtPackage.ContentPath) is not null || archive.GetEntry(OdtPackage.ManifestPath) is not null);

    public static PrivacyReport Inspect(byte[] bytes, PrivacyRequest request)
    {
        using var archive = OdtReader.OpenArchive(new MemoryStream(bytes, writable: false));
        var data = new OdtPersonalData(OdtPackage.Open(archive));
        return new PrivacyReport(
            data.Names(request.Anonymous),
            [.. data.TemplateLinks().Select(link => Uri.UnescapeDataString(link.Value))],
            [.. data.CustomProperties().Select(property => property.Attr(Ns.Meta + "name") ?? string.Empty)],
            data._package.Has(ThumbnailPath),
            data.Pictures().Count(picture => ImageMetadata.Strip(data._package.ReadBytes(picture)!) is not null),
            OdtPrivacyMark.IsSet(data.Xml(OdtPrivacyMark.SettingsPath)));
    }

    /// <returns>The clean package, and how many parts changed.</returns>
    public static (byte[] Bytes, int Changes) Scrub(byte[] bytes, PrivacyRequest request)
    {
        using var archive = OdtReader.OpenArchive(new MemoryStream(bytes, writable: false));
        var data = new OdtPersonalData(OdtPackage.Open(archive));
        var categories = request.Categories ?? [];

        if (categories.Contains(PrivacyCategory.Names)) data.AnonymizeNames(request.Anonymous);
        if (categories.Contains(PrivacyCategory.Paths)) data.Remove(OdtPackage.MetaPath, data.TemplateLinks().Select(link => link.Parent!));
        if (categories.Contains(PrivacyCategory.CustomProperties)) data.Remove(OdtPackage.MetaPath, data.CustomProperties());
        if (categories.Contains(PrivacyCategory.Thumbnail)) data.RemoveThumbnail();
        if (categories.Contains(PrivacyCategory.ImageMetadata)) data.StripPictures();

        var changes = data._edited.Count + data._replaced.Count + data._removed.Count;
        return changes == 0 ? (bytes, 0) : (data.Write(), changes);
    }

    private XDocument? Xml(string path)
    {
        if (_xml.TryGetValue(path, out var known)) return known;
        var read = _package.ReadXml(path);
        if (read is not null) _xml[path] = read;
        return read;
    }

    private IEnumerable<XElement> NamesIn(string path)
    {
        var root = Xml(path)?.Root;
        if (root is null) return [];
        var meta = root.Element(Ns.Office + "meta")?.Elements().Where(IsMetaName) ?? [];
        var authors = root.Descendants().Where(element => AuthoredElements.Contains(element.Name)).Elements(Ns.Dc + "creator");
        return meta.Concat(authors);
    }

    private static bool IsMetaName(XElement element) =>
        MetaNames.Contains(element.Name) ||
        (element.Name == Ns.Meta + "user-defined" && NamedUserDefined.Contains(element.Attr(Ns.Meta + "name")));

    private List<string> Names(string anonymous) =>
        [.. NamedParts.SelectMany(NamesIn).Select(element => element.Value.Trim())
            .Where(name => name.Length > 0 && name != anonymous).Distinct(StringComparer.Ordinal)];

    private void AnonymizeNames(string anonymous)
    {
        foreach (var path in NamedParts)
        {
            foreach (var element in NamesIn(path).Where(element => element.Value != anonymous).ToList())
            {
                // As in the DOCX: the document's own author fields go; comments and changes keep an author.
                if (element.Parent?.Name == Ns.Office + "meta") element.Remove();
                else element.Value = anonymous;
                _edited.Add(path);
            }

            if (AnonymizeInitials(path, anonymous)) _edited.Add(path);
        }

        var settings = Xml(OdtPrivacyMark.SettingsPath) ?? AddSettings();
        if (OdtPrivacyMark.Set(settings)) _edited.Add(OdtPrivacyMark.SettingsPath);
    }

    /// <summary>As in the DOCX: a comment's initials become those of the generic name.</summary>
    private bool AnonymizeInitials(string path, string anonymous)
    {
        var initial = anonymous.Length == 0 ? string.Empty : anonymous[..1];
        var changed = false;
        var initials = Xml(path)?.Root?.Descendants(Ns.Office + "annotation").Elements(Ns.Meta + "creator-initials") ?? [];
        foreach (var element in initials.Where(element => element.Value != initial).ToList())
        {
            element.Value = initial;
            changed = true;
        }

        return changed;
    }

    /// <summary>A package without <c>settings.xml</c> gets one, declared in the manifest.</summary>
    private XDocument AddSettings()
    {
        var settings = OdtPrivacyMark.Empty();
        _xml[OdtPrivacyMark.SettingsPath] = settings;
        if (Xml(OdtPackage.ManifestPath)?.Root is { } manifest)
        {
            manifest.Add(new XElement(
                Ns.Manifest + "file-entry",
                new XAttribute(Ns.Manifest + "full-path", OdtPrivacyMark.SettingsPath),
                new XAttribute(Ns.Manifest + "media-type", XmlType)));
            _edited.Add(OdtPackage.ManifestPath);
        }

        return settings;
    }

    /// <summary>The template the document was made from: a path on the author's computer.</summary>
    private List<XAttribute> TemplateLinks() =>
        [.. MetaElements(Ns.Meta + "template").Select(template => template.Attribute(Ns.XLink + "href")).OfType<XAttribute>()];

    private List<XElement> CustomProperties() =>
        [.. MetaElements(Ns.Meta + "user-defined").Where(field => !OdtProperties.OwnUserDefined.Contains(field.Attr(Ns.Meta + "name")))];

    private IEnumerable<XElement> MetaElements(XName name) =>
        Xml(OdtPackage.MetaPath)?.Root?.Element(Ns.Office + "meta")?.Elements(name) ?? [];

    private void Remove(string path, IEnumerable<XElement> elements)
    {
        foreach (var element in elements.ToList())
        {
            element.Remove();
            _edited.Add(path);
        }
    }

    private void RemoveThumbnail()
    {
        if (!_package.Has(ThumbnailPath)) return;
        _removed.Add(ThumbnailPath);
        var entries = Xml(OdtPackage.ManifestPath)?.Root?.Elements(Ns.Manifest + "file-entry")
            .Where(entry => entry.Attr(Ns.Manifest + "full-path") is ThumbnailPath or ThumbnailFolder) ?? [];
        Remove(OdtPackage.ManifestPath, entries);
    }

    /// <summary>The document's images; the thumbnail is a category of its own.</summary>
    private List<string> Pictures() =>
        [.. _package.Entries.Select(entry => entry.FullName)
            .Where(path => path.StartsWith(PicturesFolder, StringComparison.Ordinal) && !path.EndsWith('/'))];

    private void StripPictures()
    {
        foreach (var picture in Pictures())
        {
            if (ImageMetadata.Strip(_package.ReadBytes(picture)!) is { } clean) _replaced[picture] = clean;
        }
    }

    private byte[] BytesOf(string path) =>
        _edited.Contains(path) && _xml.TryGetValue(path, out var xml)
            ? XmlParts.Write(xml)
            : _replaced.GetValueOrDefault(path) ?? _package.ReadBytes(path) ?? [];

    private byte[] Write()
    {
        var addSettings = !_package.Has(OdtPrivacyMark.SettingsPath) && _edited.Contains(OdtPrivacyMark.SettingsPath);
        OdtAddedEntry[] added = addSettings
            ? [new OdtAddedEntry(OdtPrivacyMark.SettingsPath, BytesOf(OdtPrivacyMark.SettingsPath), CompressionLevel.Optimal)]
            : [];
        return OdtZip.Write(_package, path => _removed.Contains(path) ? null : BytesOf(path), added);
    }
}

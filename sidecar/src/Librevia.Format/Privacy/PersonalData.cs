using System.Text.Json.Serialization;
using System.Xml.Linq;

namespace Librevia.Format.Privacy;

public enum OfficeKind
{
    Word,
    Excel,
}

/// <summary>By content type: the Open XML ones, and the <c>vnd.ms-</c> ones of macro and template files.</summary>
internal static class OfficeContentType
{
    public static bool IsWord(string type) =>
        type.Contains("wordprocessingml", StringComparison.Ordinal) ||
        type.StartsWith("application/vnd.ms-word", StringComparison.Ordinal);

    public static bool IsExcel(string type) =>
        type.Contains("spreadsheetml", StringComparison.Ordinal) ||
        type.StartsWith("application/vnd.ms-excel", StringComparison.Ordinal);
}

/// <summary>What a DOCX or XLSX reveals about who made it, beyond its content.</summary>
public sealed record PrivacyReport(
    [property: JsonPropertyName("names")] List<string> Names,
    [property: JsonPropertyName("paths")] List<string> Paths,
    [property: JsonPropertyName("customProperties")] List<string> CustomProperties,
    [property: JsonPropertyName("thumbnail")] bool Thumbnail,
    [property: JsonPropertyName("imagesWithMetadata")] int ImagesWithMetadata,
    [property: JsonPropertyName("removePersonalInformation")] bool RemovePersonalInformation);

/// <param name="Anonymous">The name every author becomes, in the person's language.</param>
public sealed record PrivacyRequest(
    [property: JsonPropertyName("anonymous")] string Anonymous,
    [property: JsonPropertyName("categories")] List<string>? Categories = null);

public static class PrivacyCategory
{
    public const string Names = "names";
    public const string Paths = "paths";
    public const string CustomProperties = "customProperties";
    public const string Thumbnail = "thumbnail";
    public const string ImageMetadata = "imageMetadata";
}

/// <summary>
/// Inspects and cleans a package at the zip and XML level, part by part: what is not personal
/// data goes back with the same bytes.
/// </summary>
public static class PersonalData
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace AbsolutePath = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";

    private const string OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string CustomPropertiesType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    private const string ThumbnailType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    public static PrivacyReport Inspect(byte[] bytes, PrivacyRequest request)
    {
        var package = OfficePackage.Read(bytes);
        var (kind, main) = MainPartOf(package);
        return new PrivacyReport(
            PersonalNames.Find(package, kind, request.Anonymous),
            FilePaths(package, kind, main).Select(path => path.Shown).ToList(),
            CustomPropertyNames(package),
            package.RelationshipOf(string.Empty, ThumbnailType) is not null,
            Images(package).Count(part => ImageMetadata.Strip(package.Bytes(part)!) is not null),
            PrivacyFlag.IsSet(FlagHolder(package, kind, main) is { } holder ? package.Xml(holder) : null));
    }

    /// <returns>The clean package, and how many parts changed.</returns>
    public static (byte[] Bytes, int Changes) Scrub(byte[] bytes, PrivacyRequest request)
    {
        var package = OfficePackage.Read(bytes);
        var (kind, main) = MainPartOf(package);
        var categories = request.Categories ?? [];

        if (categories.Contains(PrivacyCategory.Names))
        {
            PersonalNames.Anonymize(package, kind, request.Anonymous);
            SetFlag(package, kind, main);
        }

        if (categories.Contains(PrivacyCategory.Paths)) RemoveFilePaths(package, kind, main);
        if (categories.Contains(PrivacyCategory.CustomProperties)) RemoveRootPart(package, CustomPropertiesType);
        if (categories.Contains(PrivacyCategory.Thumbnail)) RemoveRootPart(package, ThumbnailType);
        if (categories.Contains(PrivacyCategory.ImageMetadata)) StripImages(package);

        return package.Changes == 0 ? (bytes, 0) : (package.Write(), package.Changes);
    }

    private static (OfficeKind Kind, string Main) MainPartOf(OfficePackage package)
    {
        var main = package.RelationshipOf(string.Empty, OfficeDocument)?.Part;
        var type = main is null ? null : package.ContentType(main);
        if (type is not null && OfficeContentType.IsWord(type)) return (OfficeKind.Word, main!);
        if (type is not null && OfficeContentType.IsExcel(type)) return (OfficeKind.Excel, main!);

        throw new PrivacyException("sidecar.error.notOfficeDocument");
    }

    /// <summary>Word keeps the flag in <c>settings.xml</c>; Excel, in the workbook itself.</summary>
    private static string? FlagHolder(OfficePackage package, OfficeKind kind, string main) =>
        kind == OfficeKind.Excel ? main : package.RelationshipOf(main, Settings)?.Part;

    private static void SetFlag(OfficePackage package, OfficeKind kind, string main)
    {
        var holder = FlagHolder(package, kind, main) ?? AddSettings(package, main);
        if (PrivacyFlag.Set(package.Xml(holder))) package.Edited(holder);
    }

    /// <summary>A document generated outside Word may come without <c>settings.xml</c>.</summary>
    private static string AddSettings(OfficePackage package, string main)
    {
        var slash = main.LastIndexOf('/');
        var part = (slash < 0 ? string.Empty : main[..(slash + 1)]) + "settings.xml";
        package.AddXmlPart(main, part, SettingsContentType, Settings, new XElement(
            W + "settings", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName)));
        return part;
    }

    private sealed record FilePath(string Shown, Action Remove);

    /// <summary>
    /// Folders on the author's computer: the template a Word document was attached to, and the
    /// folder Excel last saved the workbook in. Both usually carry the user name.
    /// </summary>
    private static List<FilePath> FilePaths(OfficePackage package, OfficeKind kind, string main) =>
        kind == OfficeKind.Word ? TemplatePaths(package, main) : WorkbookFolders(package, main);

    private static List<FilePath> TemplatePaths(OfficePackage package, string main)
    {
        var settings = package.RelationshipOf(main, Settings)?.Part;
        var attached = settings is null ? null : package.Xml(settings)?.Root?.Element(W + "attachedTemplate");
        var id = (string?)attached?.Attribute(R + "id");
        var relationship = id is null ? null : package.RelationshipsOf(settings!).FirstOrDefault(candidate => candidate.Id == id);
        if (attached is null || relationship is null) return [];

        return
        [
            new FilePath(Uri.UnescapeDataString(relationship.Target), () =>
            {
                attached.Remove();
                package.Edited(settings!);
                package.RemoveRelationship(settings!, relationship.Id);
            }),
        ];
    }

    private static List<FilePath> WorkbookFolders(OfficePackage package, string main)
    {
        var root = package.Xml(main)?.Root;
        if (root is null) return [];

        return root.Descendants(AbsolutePath + "absPath").ToList().Select(element =>
        {
            // Excel writes it inside `mc:AlternateContent`, which would be left empty.
            var holder = element.Parent?.Name == Mc + "Choice" && element.Parent.Parent?.Name == Mc + "AlternateContent"
                ? element.Parent.Parent
                : element;
            return new FilePath((string?)element.Attribute("url") ?? string.Empty, () =>
            {
                holder.Remove();
                package.Edited(main);
            });
        }).ToList();
    }

    private static void RemoveFilePaths(OfficePackage package, OfficeKind kind, string main)
    {
        foreach (var path in FilePaths(package, kind, main)) path.Remove();
    }

    private static List<string> CustomPropertyNames(OfficePackage package)
    {
        var part = package.RelationshipOf(string.Empty, CustomPropertiesType)?.Part;
        var root = part is null ? null : package.Xml(part)?.Root;
        return root?.Elements()
            .Select(property => (string?)property.Attribute("name"))
            .OfType<string>()
            .ToList() ?? [];
    }

    private static void RemoveRootPart(OfficePackage package, string type)
    {
        var relationship = package.RelationshipOf(string.Empty, type);
        if (relationship is null) return;

        package.RemoveRelationship(string.Empty, relationship.Id);
        if (relationship.Part is not null) package.RemovePart(relationship.Part);
    }

    /// <summary>The document's images; the thumbnail is a category of its own.</summary>
    private static IEnumerable<string> Images(OfficePackage package)
    {
        var thumbnail = package.RelationshipOf(string.Empty, ThumbnailType)?.Part;
        return package.Parts
            .Where(part => part != thumbnail && package.Bytes(part) is not null)
            .Where(part => package.ContentType(part) is { } type && type.StartsWith("image/", StringComparison.Ordinal))
            .ToList();
    }

    private static void StripImages(OfficePackage package)
    {
        foreach (var part in Images(package))
        {
            if (ImageMetadata.Strip(package.Bytes(part)!) is { } clean) package.Replace(part, clean);
        }
    }
}

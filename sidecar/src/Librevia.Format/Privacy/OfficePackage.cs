using System.IO.Compression;
using System.Xml.Linq;

namespace Librevia.Format.Privacy;

/// <summary>
/// A DOCX or XLSX as its zip entries. Only the parts edited here are rewritten; every other one
/// goes back with the same bytes, in the same order.
/// </summary>
internal sealed class OfficePackage
{
    private const string ContentTypesPart = "[Content_Types].xml";
    private const string RootRelationships = "_rels/.rels";
    private static readonly XNamespace Types = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly List<string> _order = [];
    private readonly Dictionary<string, byte[]> _parts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, XDocument?> _xml = new(StringComparer.Ordinal);
    private readonly HashSet<string> _edited = new(StringComparer.Ordinal);

    public sealed record Relationship(string Id, string Type, string Target, bool External, string? Part);

    public static OfficePackage Read(byte[] bytes)
    {
        var package = new OfficePackage();
        try
        {
            using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
            foreach (var entry in archive.Entries)
            {
                using var stream = entry.Open();
                using var copy = new MemoryStream();
                stream.CopyTo(copy);
                package._order.Add(entry.FullName);
                package._parts[entry.FullName] = copy.ToArray();
            }
        }
        catch (InvalidDataException)
        {
            throw new PrivacyException("O arquivo não é um pacote do Office válido.");
        }

        if (!package._parts.ContainsKey(ContentTypesPart))
        {
            throw new PrivacyException("O arquivo não é um pacote do Office válido.");
        }

        return package;
    }

    /// <summary>How many parts changed or left.</summary>
    public int Changes { get; private set; }

    public IEnumerable<string> Parts => _order;

    public byte[]? Bytes(string part) => _parts.GetValueOrDefault(part);

    public void Replace(string part, byte[] bytes)
    {
        _parts[part] = bytes;
        Changes++;
    }

    /// <summary>Null when absent or malformed.</summary>
    public XDocument? Xml(string part)
    {
        if (_xml.TryGetValue(part, out var cached)) return cached;
        var bytes = Bytes(part);
        var parsed = bytes is null ? null : XmlParts.Read(bytes);
        _xml[part] = parsed;
        return parsed;
    }

    /// <summary>Marks the XML <see cref="Xml"/> returned as changed.</summary>
    public void Edited(string part)
    {
        if (_edited.Add(part)) Changes++;
    }

    public string? ContentType(string part)
    {
        var types = Xml(ContentTypesPart)?.Root;
        var name = "/" + part;
        var overridden = types?.Elements(Types + "Override")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("PartName"), name, StringComparison.OrdinalIgnoreCase));
        if (overridden is not null) return (string?)overridden.Attribute("ContentType");

        var extension = Path.GetExtension(part).TrimStart('.');
        return (string?)types?.Elements(Types + "Default")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase))
            ?.Attribute("ContentType");
    }

    public IEnumerable<string> PartsOfType(Func<string, bool> matches) =>
        _order.Where(part => _parts.ContainsKey(part) && ContentType(part) is { } type && matches(type)).ToList();

    /// <summary>Of a part, or of the package when <paramref name="source"/> is empty.</summary>
    public IReadOnlyList<Relationship> RelationshipsOf(string source)
    {
        var root = Xml(RelationshipsPart(source))?.Root;
        if (root is null) return [];

        return root.Elements(Relationships + "Relationship").Select(element =>
        {
            var target = (string?)element.Attribute("Target") ?? string.Empty;
            var external = string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.Ordinal);
            return new Relationship(
                (string?)element.Attribute("Id") ?? string.Empty,
                (string?)element.Attribute("Type") ?? string.Empty,
                target,
                external,
                external ? null : Resolve(source, target));
        }).ToList();
    }

    public Relationship? RelationshipOf(string source, string type) =>
        RelationshipsOf(source).FirstOrDefault(relationship => relationship.Type == type);

    public void RemoveRelationship(string source, string id)
    {
        var part = RelationshipsPart(source);
        var element = Xml(part)?.Root?.Elements(Relationships + "Relationship")
            .FirstOrDefault(candidate => (string?)candidate.Attribute("Id") == id);
        if (element is null) return;

        element.Remove();
        Edited(part);
    }

    /// <summary>With its relationships and its content type, so no reader looks for it.</summary>
    public void RemovePart(string part)
    {
        if (!_parts.Remove(part)) return;
        Changes++;
        _parts.Remove(RelationshipsPart(part));

        var types = Xml(ContentTypesPart)?.Root;
        var overridden = types?.Elements(Types + "Override")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("PartName"), "/" + part, StringComparison.OrdinalIgnoreCase));
        if (overridden is null) return;
        overridden.Remove();
        Edited(ContentTypesPart);
    }

    /// <summary>A new XML part, related from <paramref name="source"/> and with its content type.</summary>
    public XDocument AddXmlPart(string source, string part, string contentType, string relationshipType, XElement root)
    {
        var xml = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        _order.Add(part);
        _parts[part] = [];
        _xml[part] = xml;
        Edited(part);

        var types = Xml(ContentTypesPart)!.Root!;
        types.Add(new XElement(Types + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", contentType)));
        Edited(ContentTypesPart);

        var relationshipsPart = RelationshipsPart(source);
        var relationships = Xml(relationshipsPart)?.Root;
        if (relationships is null)
        {
            relationships = new XElement(Relationships + "Relationships");
            _order.Add(relationshipsPart);
            _parts[relationshipsPart] = [];
            _xml[relationshipsPart] = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), relationships);
        }

        var ids = relationships.Elements().Select(element => (string?)element.Attribute("Id")).ToHashSet(StringComparer.Ordinal);
        var id = Enumerable.Range(1, ids.Count + 1).Select(number => $"rId{number}").First(candidate => !ids.Contains(candidate));
        var slash = source.LastIndexOf('/');
        var target = slash < 0 ? part : part[(slash + 1)..];
        relationships.Add(new XElement(
            Relationships + "Relationship",
            new XAttribute("Id", id),
            new XAttribute("Type", relationshipType),
            new XAttribute("Target", target)));
        Edited(relationshipsPart);
        return xml;
    }

    public byte[] Write()
    {
        using var result = new MemoryStream();
        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var part in _order)
            {
                if (!_parts.TryGetValue(part, out var bytes)) continue;
                if (_edited.Contains(part) && _xml.GetValueOrDefault(part) is { } xml) bytes = XmlParts.Write(xml);

                using var target = output.CreateEntry(part, CompressionLevel.Optimal).Open();
                target.Write(bytes, 0, bytes.Length);
            }
        }

        return result.ToArray();
    }

    private static string RelationshipsPart(string source)
    {
        if (source.Length == 0) return RootRelationships;
        var slash = source.LastIndexOf('/');
        return slash < 0 ? $"_rels/{source}.rels" : $"{source[..slash]}/_rels/{source[(slash + 1)..]}.rels";
    }

    private static string Resolve(string source, string target)
    {
        if (target.StartsWith('/')) return target[1..];

        var slash = source.LastIndexOf('/');
        var folder = slash < 0 ? [] : source[..slash].Split('/').ToList();
        foreach (var segment in target.Split('/'))
        {
            if (segment == "..")
            {
                if (folder.Count > 0) folder.RemoveAt(folder.Count - 1);
            }
            else if (segment is not ("." or ""))
            {
                folder.Add(segment);
            }
        }

        return string.Join('/', folder);
    }
}

public sealed class PrivacyException(string message) : Exception(message);

using System.Security.Cryptography;
using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// The images of a rewritten block: one the package already has keeps its entry; a new one goes in
/// <c>Pictures/</c> under a name of its own and into the manifest. Without a package, as
/// <see cref="OdtTemplate"/> creates one, every image is new.
/// </summary>
internal sealed class OdtPictures(OdtPackage? package)
{
    private const string Folder = "Pictures/";
    private const string NamePrefix = "Librevia";

    private static readonly Dictionary<string, string> Extensions = new(StringComparer.Ordinal)
    {
        ["image/png"] = ".png",
        ["image/jpeg"] = ".jpg",
        ["image/gif"] = ".gif",
        ["image/bmp"] = ".bmp",
    };

    private Dictionary<string, string>? _existing;
    private readonly List<(string Path, string MediaType, byte[] Bytes)> _added = [];

    public IReadOnlyList<(string Path, string MediaType, byte[] Bytes)> Added => _added;

    public static bool IsSupported(string mediaType) => Extensions.ContainsKey(mediaType);

    /// <summary>The path inside the package for these bytes.</summary>
    public string PathOf(byte[] bytes, string mediaType)
    {
        var hash = Convert.ToHexString(SHA256.HashData(bytes));
        var existing = _existing ??= Existing();
        if (existing.TryGetValue(hash, out var known)) return known;

        var number = 1;
        var path = NameOf(number, mediaType);
        while (package?.Has(path) == true || _added.Any(added => added.Path == path)) path = NameOf(++number, mediaType);

        _added.Add((path, mediaType, bytes));
        existing[hash] = path;
        return path;
    }

    private static string NameOf(int number, string mediaType) => $"{Folder}{NamePrefix}{number}{Extensions[mediaType]}";

    /// <summary>The package's images by content: an unchanged image in an edited block keeps its file.</summary>
    private Dictionary<string, string> Existing()
    {
        var byHash = new Dictionary<string, string>(StringComparer.Ordinal);
        if (package is null) return byHash;
        foreach (var (path, type) in package.MediaTypes)
        {
            if (!Extensions.ContainsKey(type) || package.ReadBytes(path) is not { } bytes) continue;
            byHash.TryAdd(Convert.ToHexString(SHA256.HashData(bytes)), path);
        }

        return byHash;
    }

    /// <summary>One <c>manifest:file-entry</c> per new image.</summary>
    public IEnumerable<XElement> ManifestEntries() => _added.Select(added => new XElement(
        Ns.Manifest + "file-entry",
        new XAttribute(Ns.Manifest + "full-path", added.Path),
        new XAttribute(Ns.Manifest + "media-type", added.MediaType)));
}

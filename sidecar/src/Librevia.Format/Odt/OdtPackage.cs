using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// The ODT zip, read with the defences an untrusted file needs: every entry is decompressed up to a
/// ceiling, whatever its header claims, and XML goes through <see cref="XmlParts"/> (no DTD, no
/// resolver).
/// </summary>
internal sealed class OdtPackage
{
    public const string TextMimeType = "application/vnd.oasis.opendocument.text";
    public const string TemplateMimeType = "application/vnd.oasis.opendocument.text-template";

    /// <summary>LibreOffice writes it in an <c>.odt</c> saved from HTML; the content is ODF text.</summary>
    private const string WebTextMimeType = "application/vnd.oasis.opendocument.text-web";
    private const string FormulaMimeType = "application/vnd.oasis.opendocument.formula";

    /// <summary>Against a zip bomb: no real part of a text document comes near it.</summary>
    public const long MaxEntryBytes = 64L * 1024 * 1024;

    /// <summary>The whole package, decompressed, across every entry read.</summary>
    public const long MaxTotalBytes = 256L * 1024 * 1024;

    private const int CopyBufferBytes = 81920;

    public const string ContentPath = "content.xml";
    public const string StylesPath = "styles.xml";
    public const string MetaPath = "meta.xml";
    public const string ManifestPath = "META-INF/manifest.xml";
    public const string MimeTypePath = "mimetype";

    private readonly Dictionary<string, ZipArchiveEntry> _entries;
    private readonly List<ZipArchiveEntry> _ordered;

    /// <summary>Each entry is decompressed once: the writer copies what the reader already read.</summary>
    private readonly Dictionary<string, byte[]> _read = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _mediaTypes = new(StringComparer.Ordinal);
    private long _budget = MaxTotalBytes;

    public XDocument Content { get; }
    public XDocument? Styles { get; }
    public XDocument? Meta { get; }
    public XDocument? Manifest { get; }

    /// <summary>As <c>mimetype</c> declares it, or the manifest's root entry; <c>null</c> when neither does.</summary>
    public string? MimeType { get; private set; }

    private OdtPackage(ZipArchive archive)
    {
        _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
        _ordered = [];
        foreach (var entry in archive.Entries)
        {
            if (_entries.TryAdd(entry.FullName, entry)) _ordered.Add(entry);
        }

        var manifest = ReadXml(ManifestPath);
        Manifest = manifest;
        if (manifest is not null) ReadManifest(manifest);
        CheckMimeType();

        Content = ReadXml(ContentPath) ?? throw new OdtException(OdtException.WithoutContent);
        Styles = ReadXml(StylesPath);
        Meta = ReadXml(MetaPath);
    }

    public static OdtPackage Open(ZipArchive archive) => new(archive);

    public IReadOnlyList<ZipArchiveEntry> Entries => _ordered;

    public bool Has(string path) => _entries.ContainsKey(path);

    public byte[]? ReadBytes(string path)
    {
        if (_read.TryGetValue(path, out var known)) return known;
        if (!_entries.TryGetValue(path, out var entry)) return null;
        if (entry.Length > MaxEntryBytes || entry.Length > _budget) throw new OdtException(OdtException.TooLarge);

        try
        {
            var bytes = Decompress(entry);
            _budget -= bytes.Length;
            _read[path] = bytes;
            return bytes;
        }
        catch (Exception problem) when (problem is InvalidDataException or IOException or NotSupportedException)
        {
            throw new OdtException(OdtException.Unreadable, problem);
        }
    }

    /// <summary>The header's size is the file's word: the ceiling is checked on what comes out.</summary>
    private byte[] Decompress(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var buffer = new MemoryStream();
        var chunk = new byte[CopyBufferBytes];
        int read;
        while ((read = stream.Read(chunk, 0, chunk.Length)) > 0)
        {
            if (buffer.Length + read > MaxEntryBytes || buffer.Length + read > _budget)
            {
                throw new OdtException(OdtException.TooLarge);
            }

            buffer.Write(chunk, 0, read);
        }

        return buffer.ToArray();
    }

    public IReadOnlyDictionary<string, string> MediaTypes => _mediaTypes;

    public string? MediaTypeOf(string path) => _mediaTypes.GetValueOrDefault(path);

    /// <summary>An embedded object (<c>Object 1</c>) that is a formula.</summary>
    public bool IsFormula(string objectPath) =>
        string.Equals(MediaTypeOf(objectPath.TrimEnd('/') + "/"), FormulaMimeType, StringComparison.Ordinal);

    public XDocument? ReadXml(string path)
    {
        var bytes = ReadBytes(path);
        if (bytes is null) return null;
        return XmlParts.Read(bytes) ?? throw new OdtException(OdtException.Unreadable);
    }

    private void ReadManifest(XDocument manifest)
    {
        foreach (var entry in manifest.Descendants(Ns.Manifest + "file-entry"))
        {
            // The key is per part: the package would need a password we do not ask for.
            if (entry.Element(Ns.Manifest + "encryption-data") is not null)
            {
                throw new OdtException(OdtException.Encrypted);
            }

            if (entry.Attr(Ns.Manifest + "full-path") is { } path && entry.Attr(Ns.Manifest + "media-type") is { } type)
            {
                _mediaTypes[path] = type;
            }
        }
    }

    /// <summary>The <c>mimetype</c> entry, or the manifest's root entry when it is missing.</summary>
    private void CheckMimeType()
    {
        var declared = ReadBytes(MimeTypePath) is { } bytes ? MimeTypeIn(bytes) : MediaTypeOf("/");
        if (declared is not null && !IsText(declared)) throw new OdtException(OdtException.NotText);
        MimeType = declared;
    }

    public static string MimeTypeIn(byte[] mimeTypeEntry) => Encoding.ASCII.GetString(mimeTypeEntry).Trim();

    public static bool IsText(string mimeType) => mimeType is TextMimeType or TemplateMimeType or WebTextMimeType;
}

/// <summary>Fails with the catalog key of a sentence ready for the user.</summary>
public sealed class OdtException(string key, Exception? inner = null) : Exception(key, inner)
{
    public const string Unreadable = "sidecar.error.odtUnreadable";
    public const string WithoutContent = "sidecar.error.odtWithoutContent";
    public const string TooLarge = "sidecar.error.odtTooLarge";
    public const string Encrypted = "sidecar.error.odtEncrypted";
    public const string NotText = "sidecar.error.odtNotText";
}

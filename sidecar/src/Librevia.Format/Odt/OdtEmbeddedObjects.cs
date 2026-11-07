using System.Text;
using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// An embedded object (<c>Object 1/</c>, an equation or a chart) whose last frame went with an
/// edited block: its folder, its replacement picture and their manifest entries go too. Only one the
/// original body used, and that neither the new body nor <c>styles.xml</c> points to any more.
/// </summary>
internal sealed class OdtEmbeddedObjects
{
    private const string ObjectMediaPrefix = "application/vnd.oasis.opendocument.";
    private const string Replacements = "ObjectReplacements/";
    private const string Here = "./";

    private readonly HashSet<string> _orphans;
    private readonly HashSet<string> _referenced;

    private OdtEmbeddedObjects(HashSet<string> orphans, HashSet<string> referenced)
    {
        _orphans = orphans;
        _referenced = referenced;
    }

    public bool IsEmpty => _orphans.Count == 0;

    public static OdtEmbeddedObjects Find(OdtPackage package, string content, XDocument? styles)
    {
        var referenced = References(XmlParts.Read(Encoding.UTF8.GetBytes(content)));
        referenced.UnionWith(References(styles));
        var before = References(package.Content);
        var orphans = package.MediaTypes
            .Where(entry => entry.Value.StartsWith(ObjectMediaPrefix, StringComparison.Ordinal) && IsTopFolder(entry.Key))
            .Select(entry => entry.Key.TrimEnd('/'))
            .Where(folder => before.Contains(folder) && !referenced.Contains(folder))
            .ToHashSet(StringComparer.Ordinal);
        return new OdtEmbeddedObjects(orphans, referenced);
    }

    /// <summary>A part or a manifest path (<c>Object 1/</c>, <c>Object 1/content.xml</c>) that goes.</summary>
    public bool Drops(string path) =>
        _orphans.Any(folder => path.StartsWith(folder + "/", StringComparison.Ordinal) ||
                               (path == Replacements + folder && !_referenced.Contains(path)));

    private static bool IsTopFolder(string path) =>
        path.Length > 1 && path.EndsWith('/') && path.IndexOf('/') == path.Length - 1;

    private static HashSet<string> References(XDocument? document) =>
        (document?.Descendants() ?? [])
        .Select(element => element.Attr(Ns.XLink + "href"))
        .OfType<string>()
        .Select(Normalized)
        .ToHashSet(StringComparer.Ordinal);

    private static string Normalized(string href)
    {
        var path = href.StartsWith(Here, StringComparison.Ordinal) ? href[Here.Length..] : href;
        return path.TrimEnd('/');
    }
}

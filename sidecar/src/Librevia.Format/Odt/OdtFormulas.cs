using System.Text;
using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// Each new equation is an embedded formula object, a folder with its MathML, as LibreOffice Math
/// writes it. Without the replacement image (<c>ObjectReplacements/</c>): whoever opens the file
/// draws the formula.
/// </summary>
internal sealed class OdtFormulas(Func<string, bool> taken)
{
    public const string MediaType = "application/vnd.oasis.opendocument.formula";

    private const string FolderPrefix = "Object ";
    private const string ContentName = "/content.xml";
    private const string ObjectVersion = OdtTemplate.OdfVersion;

    private readonly List<(string Folder, byte[] Content)> _added = [];

    public IReadOnlyList<(string Folder, byte[] Content)> Added => _added;

    /// <returns>The folder, <c>Object 1</c>, without the trailing slash.</returns>
    public string Add(XElement math, bool display)
    {
        var number = 1;
        string folder;
        while (taken(folder = FolderPrefix + number) || taken(folder + ContentName) || _added.Any(added => added.Folder == folder)) number++;

        var tree = new XElement(math.Name, new XAttribute("xmlns", math.Name.NamespaceName), math.Attributes(), math.Nodes());
        tree.SetAttributeValue("display", display ? "block" : "inline");
        var text = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + tree.ToString(SaveOptions.DisableFormatting);
        _added.Add((folder, Encoding.UTF8.GetBytes(text)));
        return folder;
    }

    public static string ContentPathOf(string folder) => folder + ContentName;

    /// <summary>The folder as a formula document, and its content.</summary>
    public IEnumerable<XElement> ManifestEntries() => _added.SelectMany(added => new[]
    {
        new XElement(
            Ns.Manifest + "file-entry",
            new XAttribute(Ns.Manifest + "full-path", added.Folder + "/"),
            new XAttribute(Ns.Manifest + "version", ObjectVersion),
            new XAttribute(Ns.Manifest + "media-type", MediaType)),
        new XElement(
            Ns.Manifest + "file-entry",
            new XAttribute(Ns.Manifest + "full-path", ContentPathOf(added.Folder)),
            new XAttribute(Ns.Manifest + "media-type", "text/xml")),
    });
}

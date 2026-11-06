using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// The automatic styles a save creates, appended to the part's <c>office:automatic-styles</c>.
/// Names never collide with a style the package already has, and two equal styles are one.
/// </summary>
internal sealed class OdtAutomaticStyles(IEnumerable<string> taken)
{
    public const string Paragraph = "LvP";
    public const string Text = "LvT";
    public const string Table = "LvTable";
    public const string Column = "LvColumn";
    public const string Cell = "LvCell";
    public const string List = "LvL";
    public const string Band = "LvMP";
    public const string Graphic = "LvFr";
    public const string Section = "LvSect";
    public const string PageLayout = "LvPm";

    public const string GraphicFamily = "graphic";
    public const string SectionFamily = "section";

    private readonly HashSet<string> _taken = new(taken, StringComparer.Ordinal);
    private readonly Dictionary<string, string> _bySignature = new(StringComparer.Ordinal);
    private readonly List<XElement> _created = [];

    public IReadOnlyList<XElement> Created => _created;

    /// <summary>Every <c>style:name</c> in the parts, automatic or named: one namespace for safety.</summary>
    public static IEnumerable<string> NamesIn(params XDocument?[] parts) =>
        parts.SelectMany(part => part?.Descendants() ?? [])
            .Select(element => element.Attr(Ns.Style + "name"))
            .OfType<string>();

    /// <param name="style">Without <c>style:name</c>: it gets one here.</param>
    public string Add(string prefix, XElement style)
    {
        var signature = prefix + style.ToString(SaveOptions.DisableFormatting);
        if (_bySignature.TryGetValue(signature, out var known)) return known;

        var number = 1;
        string name;
        while (!_taken.Add(name = $"{prefix}{number}")) number++;

        _created.Add(new XElement(style.Name, new XAttribute(Ns.Style + "name", name), style.Attributes(), style.Nodes()));
        _bySignature[signature] = name;
        return name;
    }

    public string Style(string prefix, string family, string? parent, params XElement?[] properties) =>
        Style(prefix, family, parent, masterPage: null, properties);

    /// <param name="masterPage">The master page the paragraph or table opens.</param>
    public string Style(string prefix, string family, string? parent, string? masterPage, params XElement?[] properties)
    {
        var style = new XElement(
            Ns.Style + "style",
            new XAttribute(Ns.Style + "family", family),
            parent is null ? null : new XAttribute(Ns.Style + "parent-style-name", parent),
            masterPage is null ? null : new XAttribute(Ns.Style + "master-page-name", masterPage),
            properties.Where(element => element is not null && (element.HasAttributes || element.HasElements)));
        return Add(prefix, style);
    }
}

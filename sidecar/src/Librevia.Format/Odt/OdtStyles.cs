using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>A <c>style:style</c>, named (<c>office:styles</c>) or automatic.</summary>
internal sealed record OdtStyle(string Name, string Family, XElement Element, bool Automatic)
{
    /// <summary>The editor's id: Word's for a LibreOffice built-in (<see cref="OdtWordStyles"/>).</summary>
    public string Id => OdtWordStyles.IdOf(Family, Name);

    public string? Parent => Element.Attr(Ns.Style + "parent-style-name");
    public string? DisplayName => Element.Attr(Ns.Style + "display-name");
    public string? Next => Element.Attr(Ns.Style + "next-style-name");
    public string? ListStyle => Element.Attr(Ns.Style + "list-style-name");
    public string? MasterPage => Element.Attr(Ns.Style + "master-page-name");
    public int? OutlineLevel => OdtLength.Integer(Element.Attr(Ns.Style + "default-outline-level"));

    public XElement? Properties(string kind) => Element.Element(Ns.Style + kind);
}

/// <summary>
/// The styles <c>content.xml</c> sees: its own automatic styles and the named ones of
/// <c>styles.xml</c>. Resolving walks the parent chain from the family's default style.
/// </summary>
internal sealed class OdtStyles
{
    public const string ParagraphFamily = "paragraph";
    public const string TextFamily = "text";
    public const string DefaultParagraphStyle = "Standard";

    /// <summary>The line a border, a rule or a header's rule draws when the model gives no other.</summary>
    public const string DefaultBorder = "0.5pt solid #000000";

    /// <summary>A loop in the parent chain is a broken file, not an infinite walk.</summary>
    private const int MaxChain = 32;

    private readonly Dictionary<(string Family, string Name), OdtStyle> _named;
    private readonly Dictionary<(string Family, string Name), OdtStyle> _automatic = [];
    private readonly Dictionary<string, XElement> _defaults;
    private readonly Dictionary<string, XElement> _listStyles;

    public OdtFonts Fonts { get; }

    private OdtStyles(
        Dictionary<(string, string), OdtStyle> named,
        Dictionary<string, XElement> defaults,
        Dictionary<string, XElement> listStyles,
        OdtFonts fonts)
    {
        _named = named;
        _defaults = defaults;
        _listStyles = new Dictionary<string, XElement>(listStyles, StringComparer.Ordinal);
        Fonts = fonts;
    }

    public static OdtStyles Read(XDocument content, XDocument? styles)
    {
        var fonts = new OdtFonts(styles, content);
        var office = styles?.Root?.Element(Ns.Office + "styles");
        var named = new Dictionary<(string, string), OdtStyle>();
        foreach (var style in StylesIn(office, automatic: false)) named.TryAdd((style.Family, style.Name), style);

        var defaults = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var element in office?.Elements(Ns.Style + "default-style") ?? [])
        {
            if (element.Attr(Ns.Style + "family") is { } family) defaults.TryAdd(family, element);
        }

        var read = new OdtStyles(named, defaults, ListStylesIn(office), fonts);
        read.AddAutomatic(content.Root?.Element(Ns.Office + "automatic-styles"));
        return read;
    }

    public IEnumerable<OdtStyle> Named => _named.Values;

    public bool IsNamed(string family, string name) => _named.ContainsKey((family, name));

    public OdtStyle? Find(string family, string? name)
    {
        if (string.IsNullOrEmpty(name)) return null;
        return _automatic.GetValueOrDefault((family, name)) ?? _named.GetValueOrDefault((family, name));
    }

    private XElement? DefaultStyle(string family) => _defaults.GetValueOrDefault(family);

    public XElement? ListStyle(string? name) =>
        string.IsNullOrEmpty(name) ? null : _listStyles.GetValueOrDefault(name);

    /// <summary>The named style under an automatic one, or the style itself.</summary>
    public OdtStyle? NamedOf(string family, string? name)
    {
        var style = Find(family, name);
        return style is { Automatic: true } ? NamedParent(style) : style;
    }

    /// <summary>From the root to <paramref name="name"/>, defaults not included.</summary>
    public List<OdtStyle> Chain(string family, string? name)
    {
        var chain = new List<OdtStyle>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var style = Find(family, name);
        while (style is not null && chain.Count < MaxChain && seen.Add(style.Name))
        {
            chain.Insert(0, style);
            style = NamedParent(style);
        }

        return chain;
    }

    /// <summary>A parent is always a named style, even an automatic style's.</summary>
    private OdtStyle? NamedParent(OdtStyle style) =>
        string.IsNullOrEmpty(style.Parent) ? null : _named.GetValueOrDefault((style.Family, style.Parent));

    public ParagraphFormat DefaultParagraph() =>
        ParagraphFormat.Of(DefaultStyle(ParagraphFamily)?.Element(Ns.Style + "paragraph-properties"));

    public CharacterFormat DefaultCharacter() =>
        CharacterFormat.Of(DefaultStyle(ParagraphFamily)?.Element(Ns.Style + "text-properties"), Fonts);

    public ParagraphFormat Paragraph(IEnumerable<OdtStyle> chain) =>
        chain.Aggregate(ParagraphFormat.Empty, (format, style) =>
            format.Over(ParagraphFormat.Of(style.Properties("paragraph-properties"))));

    public CharacterFormat Character(IEnumerable<OdtStyle> chain, CharacterFormat under) =>
        chain.Aggregate(under, (format, style) =>
            format.Over(CharacterFormat.Of(style.Properties("text-properties"), Fonts)));

    /// <summary>The list style a paragraph style names, walking up to the named one.</summary>
    public string? ListStyleOfParagraph(string? name)
    {
        var chain = Chain(ParagraphFamily, name);
        for (var index = chain.Count - 1; index >= 0; index--)
        {
            if (chain[index].ListStyle is { Length: > 0 } list) return list;
        }

        return null;
    }

    private void AddAutomatic(XElement? automatic)
    {
        foreach (var style in StylesIn(automatic, automatic: true)) _automatic.TryAdd((style.Family, style.Name), style);
        foreach (var (name, list) in ListStylesIn(automatic)) _listStyles[name] = list;
    }

    private static IEnumerable<OdtStyle> StylesIn(XElement? container, bool automatic)
    {
        foreach (var element in container?.Elements(Ns.Style + "style") ?? [])
        {
            if (element.Attr(Ns.Style + "name") is { Length: > 0 } name &&
                element.Attr(Ns.Style + "family") is { Length: > 0 } family)
            {
                yield return new OdtStyle(name, family, element, automatic);
            }
        }
    }

    private static Dictionary<string, XElement> ListStylesIn(XElement? container)
    {
        var lists = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var element in container?.Elements(Ns.Text + "list-style") ?? [])
        {
            if (element.Attr(Ns.Style + "name") is { Length: > 0 } name) lists.TryAdd(name, element);
        }

        return lists;
    }
}

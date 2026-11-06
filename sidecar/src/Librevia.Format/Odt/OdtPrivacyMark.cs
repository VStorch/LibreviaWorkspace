using System.Xml.Linq;

namespace Librevia.Format.Odt;

/// <summary>
/// "Remove personal information on save" for an ODT. ODF has no such setting per document, unlike
/// Word's <c>w:removePersonalInformation</c>; so the inspector leaves its own item in
/// <c>settings.xml</c>, which LibreOffice ignores, and the reader turns it into the same property.
/// </summary>
internal static class OdtPrivacyMark
{
    public const string SettingsPath = "settings.xml";

    private static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    private const string ItemSet = "ooo:configuration-settings";
    private const string Item = "LibreviaRemovePersonalInformation";
    private const string BooleanType = "boolean";
    private const string True = "true";

    public static bool IsSet(XDocument? settings) =>
        ItemSetOf(settings)?.Elements(Config + "config-item").Any(item =>
            item.Attr(Config + "name") == Item && item.Value.Trim() == True) == true;

    /// <returns>Whether the XML changed.</returns>
    public static bool Set(XDocument settings)
    {
        if (IsSet(settings) || settings.Root is not { } root) return false;

        var items = ItemSetOf(settings);
        if (items is null)
        {
            var holder = root.Element(Ns.Office + "settings") ?? new XElement(Ns.Office + "settings");
            if (holder.Parent is null) root.Add(holder);
            items = new XElement(Config + "config-item-set", new XAttribute(Config + "name", ItemSet));
            holder.Add(items);
        }

        items.Elements(Config + "config-item").Where(item => item.Attr(Config + "name") == Item).Remove();
        items.Add(new XElement(
            Config + "config-item",
            new XAttribute(Config + "name", Item),
            new XAttribute(Config + "type", BooleanType),
            True));
        return true;
    }

    /// <summary>For a package that came without <c>settings.xml</c>.</summary>
    public static XDocument Empty() => new(new XElement(
        Ns.Office + "document-settings",
        new XAttribute(XNamespace.Xmlns + "office", Ns.Office.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "config", Config.NamespaceName),
        new XAttribute(Ns.Office + "version", "1.3"),
        new XElement(Ns.Office + "settings")));

    private static XElement? ItemSetOf(XDocument? settings) =>
        settings?.Root?.Element(Ns.Office + "settings")?.Elements(Config + "config-item-set")
            .FirstOrDefault(set => set.Attr(Config + "name") == ItemSet);
}

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

    private const string ItemSet = "ooo:configuration-settings";
    private const string Item = "LibreviaRemovePersonalInformation";
    private const string BooleanType = "boolean";
    private const string True = "true";

    public static bool IsSet(XDocument? settings) =>
        ItemSetOf(settings)?.Elements(Ns.Config + "config-item").Any(item =>
            item.Attr(Ns.Config + "name") == Item && item.Value.Trim() == True) == true;

    /// <returns>Whether the XML changed.</returns>
    public static bool Set(XDocument settings)
    {
        if (IsSet(settings) || settings.Root is not { } root) return false;

        var items = ItemSetOf(settings);
        if (items is null)
        {
            var holder = root.Element(Ns.Office + "settings") ?? new XElement(Ns.Office + "settings");
            if (holder.Parent is null) root.Add(holder);
            items = new XElement(Ns.Config + "config-item-set", new XAttribute(Ns.Config + "name", ItemSet));
            holder.Add(items);
        }

        items.Elements(Ns.Config + "config-item").Where(item => item.Attr(Ns.Config + "name") == Item).Remove();
        items.Add(new XElement(
            Ns.Config + "config-item",
            new XAttribute(Ns.Config + "name", Item),
            new XAttribute(Ns.Config + "type", BooleanType),
            True));
        return true;
    }

    /// <summary>For a package that came without <c>settings.xml</c>.</summary>
    public static XDocument Empty() => new(new XElement(
        Ns.Office + "document-settings",
        new XAttribute(XNamespace.Xmlns + "office", Ns.Office.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "config", Ns.Config.NamespaceName),
        new XAttribute(Ns.Office + "version", OdtTemplate.OdfVersion),
        new XElement(Ns.Office + "settings")));

    private static XElement? ItemSetOf(XDocument? settings) =>
        settings?.Root?.Element(Ns.Office + "settings")?.Elements(Ns.Config + "config-item-set")
            .FirstOrDefault(set => set.Attr(Ns.Config + "name") == ItemSet);
}

using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>meta.xml</c>, only when a property differs from the file's: the inverse of
/// <see cref="OdtProperties"/>. Each field is a patch, as in <see cref="DocumentProperties"/>: null
/// leaves it, empty clears it.
/// </summary>
internal static class OdtMetaPart
{
    private const string KeywordSeparator = ";";

    private static readonly (XName Name, Func<PropertiesDto, string?> Get)[] Fields =
    [
        (Ns.Dc + "title", properties => properties.Title),
        (Ns.Dc + "description", properties => properties.Description),
        (Ns.Dc + "subject", properties => properties.Subject),
        (Ns.Meta + "initial-creator", properties => properties.Creator),
        (Ns.Dc + "creator", properties => properties.LastModifiedBy),
        (Ns.Meta + "creation-date", properties => properties.Created),
        (Ns.Dc + "date", properties => properties.Modified),
        (Ns.Dc + "language", properties => properties.Language),
        (Ns.Meta + "editing-cycles", properties => properties.Revision),
    ];

    private static readonly (string Name, Func<PropertiesDto, string?> Get)[] UserDefined =
    [
        ("Category", properties => properties.Category),
        ("Company", properties => properties.Company),
        ("Manager", properties => properties.Manager),
    ];

    /// <returns>The new part, or <c>null</c> when the file already says it.</returns>
    public static byte[]? Write(XDocument? meta, PropertiesDto? known, PropertiesDto? wanted)
    {
        if (wanted is null) return null;
        known ??= new PropertiesDto();

        var document = meta is null ? Empty() : new XDocument(meta);
        var office = document.Root!.Element(Ns.Office + "meta") ?? AddOffice(document.Root!);
        var changed = false;
        foreach (var (name, get) in Fields)
        {
            if (get(wanted) is not { } value || value == (get(known) ?? string.Empty)) continue;
            office.Elements(name).Remove();
            if (value.Length > 0) office.Add(new XElement(name, value));
            changed = true;
        }

        changed |= Keywords(office, known.Keywords, wanted.Keywords);
        foreach (var (field, get) in UserDefined) changed |= Field(office, field, get(known), get(wanted));
        return changed ? XmlParts.Write(document) : null;
    }

    private static bool Keywords(XElement office, string? known, string? wanted)
    {
        if (wanted is null || wanted == (known ?? string.Empty)) return false;
        office.Elements(Ns.Meta + "keyword").Remove();
        foreach (var keyword in wanted.Split(KeywordSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            office.Add(new XElement(Ns.Meta + "keyword", keyword));
        }

        return true;
    }

    private static bool Field(XElement office, string field, string? known, string? wanted)
    {
        if (wanted is null || wanted == (known ?? string.Empty)) return false;
        office.Elements(Ns.Meta + "user-defined").Where(element => element.Attr(Ns.Meta + "name") == field).Remove();
        if (wanted.Length > 0) office.Add(new XElement(Ns.Meta + "user-defined", new XAttribute(Ns.Meta + "name", field), wanted));
        return true;
    }

    private static XElement AddOffice(XElement root)
    {
        var office = new XElement(Ns.Office + "meta");
        root.Add(office);
        return office;
    }

    /// <summary>A package without <c>meta.xml</c> gets one when a property is set.</summary>
    private static XDocument Empty() => new(new XElement(
        Ns.Office + "document-meta",
        new XAttribute(XNamespace.Xmlns + "office", Ns.Office.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "dc", Ns.Dc.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "meta", Ns.Meta.NamespaceName),
        new XAttribute(Ns.Office + "version", "1.3"),
        new XElement(Ns.Office + "meta")));
}

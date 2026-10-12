using System.Xml.Linq;

namespace Librevia.Format.Privacy;

/// <summary>
/// "Remove personal information from this file on save", as Word and Excel keep it: Word's
/// <c>w:removePersonalInformation</c> in <c>settings.xml</c>, Excel's <c>filterPrivacy</c> in
/// <c>workbook.xml</c>. With it, Word signs new comments and changes as "Author", and so does this
/// editor.
/// </summary>
internal static class PrivacyFlag
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>The children of <c>w:settings</c> that come before the flag, in schema order.</summary>
    private static readonly XName[] BeforeInSettings = [W + "writeProtection", W + "view", W + "zoom"];

    private static readonly XName[] BeforeInWorkbook = [S + "fileVersion", S + "fileSharing"];

    /// <param name="xml">Word's <c>settings.xml</c> or Excel's <c>workbook.xml</c>.</param>
    public static bool IsSet(XDocument? xml)
    {
        var root = xml?.Root;
        if (root is null) return false;
        if (root.Name == W + "settings") return root.Element(W + "removePersonalInformation") is not null;

        var value = (string?)root.Element(S + "workbookPr")?.Attribute("filterPrivacy");
        return value is "1" or "true";
    }

    /// <returns>Whether the XML changed.</returns>
    public static bool Set(XDocument? xml)
    {
        var root = xml?.Root;
        if (root is null || IsSet(xml)) return false;

        if (root.Name == W + "settings")
        {
            InsertAfter(root, BeforeInSettings, new XElement(W + "removePersonalInformation"));
            return true;
        }

        if (root.Name != S + "workbook") return false;

        var properties = root.Element(S + "workbookPr");
        if (properties is null)
        {
            properties = new XElement(S + "workbookPr");
            InsertAfter(root, BeforeInWorkbook, properties);
        }

        properties.SetAttributeValue("filterPrivacy", "1");
        return true;
    }

    private static void InsertAfter(XElement root, XName[] before, XElement element)
    {
        var last = root.Elements().LastOrDefault(child => before.Contains(child.Name));
        if (last is not null) last.AddAfterSelf(element);
        else root.AddFirst(element);
    }
}

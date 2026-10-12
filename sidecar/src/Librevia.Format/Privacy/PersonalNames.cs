using System.Xml.Linq;

namespace Librevia.Format.Privacy;

/// <summary>
/// The names a file carries: authorship in <c>docProps/</c>, and the author of each comment and
/// tracked change. Removing them works as in Word: the properties go, and every author becomes the
/// same generic name.
/// </summary>
internal static class PersonalNames
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private const string CoreType = "application/vnd.openxmlformats-package.core-properties+xml";
    private const string AppType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    private const string SheetCommentsType = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    private const string PersonsType = "application/vnd.ms-excel.person+xml";

    private static readonly XName[] PropertyNames =
    [
        Dc + "creator", Cp + "lastModifiedBy", Ep + "Company", Ep + "Manager",
    ];

    private static readonly XName[] AuthorAttributes = [W + "author", W15 + "author"];

    public static List<string> Find(OfficePackage package, OfficeKind kind, string anonymous)
    {
        var names = new List<string>();
        void Add(string? name)
        {
            var trimmed = name?.Trim();
            if (string.IsNullOrEmpty(trimmed) || trimmed == anonymous || names.Contains(trimmed)) return;
            names.Add(trimmed);
        }

        foreach (var root in Roots(package, PartsOfType(package, CoreType, AppType)))
        {
            foreach (var name in PropertiesIn(root).SelectMany(element => element.Value.Split(';'))) Add(name);
        }

        foreach (var root in Roots(package, WordParts(package, kind)))
        {
            foreach (var attribute in AuthorsIn(root)) Add(attribute.Value);
        }

        foreach (var root in Roots(package, PartsOfType(package, SheetCommentsType)))
        {
            foreach (var author in SheetAuthorsIn(root)) Add(author.Value);
        }

        foreach (var root in Roots(package, PartsOfType(package, PersonsType)))
        {
            foreach (var person in PersonsIn(root)) Add((string?)person.Attribute("displayName"));
        }

        return names;
    }

    public static void Anonymize(OfficePackage package, OfficeKind kind, string anonymous)
    {
        foreach (var part in PartsOfType(package, CoreType, AppType)) Edit(package, part, RemoveProperties);
        foreach (var part in WordParts(package, kind)) Edit(package, part, root => AnonymizeWordPart(root, anonymous));
        foreach (var part in PartsOfType(package, SheetCommentsType))
        {
            Edit(package, part, root => AnonymizeSheetComments(root, anonymous));
        }

        foreach (var part in PartsOfType(package, PersonsType)) Edit(package, part, root => AnonymizePersons(root, anonymous));
    }

    /// <summary>Only a part that changed is rewritten.</summary>
    private static void Edit(OfficePackage package, string part, Func<XElement, bool> change)
    {
        var root = package.Xml(part)?.Root;
        if (root is not null && change(root)) package.Edited(part);
    }

    private static IEnumerable<XElement> Roots(OfficePackage package, IEnumerable<string> parts) =>
        parts.Select(part => package.Xml(part)?.Root).OfType<XElement>();

    private static IEnumerable<XElement> PropertiesIn(XElement root) =>
        root.Elements().Where(element => PropertyNames.Contains(element.Name));

    private static bool RemoveProperties(XElement root)
    {
        var found = PropertiesIn(root).ToList();
        foreach (var element in found) element.Remove();
        return found.Count > 0;
    }

    /// <summary>
    /// Body, headers, footers, notes, comments and <c>people.xml</c>: the parts that name who
    /// wrote what. A part without the word "author" is not even parsed.
    /// </summary>
    private static IEnumerable<string> WordParts(OfficePackage package, OfficeKind kind)
    {
        if (kind != OfficeKind.Word) return [];
        return package.PartsOfType(OfficeContentType.IsWord)
            .Where(part => package.Bytes(part) is { } bytes && bytes.AsSpan().IndexOf("author"u8) >= 0)
            .ToList();
    }

    private static IEnumerable<XAttribute> AuthorsIn(XElement root) =>
        root.DescendantsAndSelf().SelectMany(element => element.Attributes())
            .Where(attribute => AuthorAttributes.Contains(attribute.Name));

    private static bool AnonymizeWordPart(XElement root, string anonymous)
    {
        var changed = false;
        foreach (var attribute in AuthorsIn(root).Where(attribute => attribute.Value != anonymous).ToList())
        {
            attribute.Value = anonymous;
            changed = true;
        }

        var initial = anonymous.Length == 0 ? string.Empty : anonymous[..1];
        foreach (var comment in root.Descendants(W + "comment"))
        {
            var initials = comment.Attribute(W + "initials");
            if (initials is null || initials.Value == initial) continue;
            initials.Value = initial;
            changed = true;
        }

        // Who the person is in Microsoft's directory, beside the name.
        var presence = root.Descendants(W15 + "presenceInfo").ToList();
        foreach (var element in presence) element.Remove();
        return changed || presence.Count > 0;
    }

    private static IEnumerable<XElement> SheetAuthorsIn(XElement root) =>
        root.Element(S + "authors")?.Elements(S + "author") ?? [];

    /// <summary>Excel also writes the author at the start of the note's text, as "Name:" in bold.</summary>
    private static bool AnonymizeSheetComments(XElement root, string anonymous)
    {
        var authors = SheetAuthorsIn(root).Where(author => author.Value != anonymous).ToList();
        if (authors.Count == 0) return false;

        var names = authors.Select(author => author.Value).Where(name => name.Length > 0).ToList();
        foreach (var author in authors) author.Value = anonymous;

        foreach (var comment in root.Descendants(S + "comment"))
        {
            var text = comment.Descendants(S + "t").FirstOrDefault();
            var name = names.FirstOrDefault(name => text?.Value.StartsWith(name + ":", StringComparison.Ordinal) == true);
            if (text is not null && name is not null) text.Value = anonymous + text.Value[name.Length..];
        }

        return true;
    }

    private static IEnumerable<XElement> PersonsIn(XElement root) =>
        root.Elements().Where(element => element.Name.LocalName == "person");

    private static bool AnonymizePersons(XElement root, string anonymous)
    {
        var persons = PersonsIn(root).ToList();
        foreach (var person in persons)
        {
            person.SetAttributeValue("displayName", anonymous);
            person.SetAttributeValue("userId", null);
            person.SetAttributeValue("providerId", null);
        }

        return persons.Count > 0;
    }

    private static IEnumerable<string> PartsOfType(OfficePackage package, params string[] types) =>
        package.PartsOfType(type => types.Contains(type, StringComparer.Ordinal));
}

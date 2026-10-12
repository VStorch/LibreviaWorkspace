using System.Globalization;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Librevia.Format.Docx;

/// <summary>
/// One field of File → Properties → Custom, as Word and Excel keep it in <c>docProps/custom.xml</c>.
/// <c>Type</c> is <c>text</c>, <c>number</c>, <c>date</c> or <c>boolean</c>; <c>Value</c> is text in
/// each case: a number with a dot, a date as the file's instant, a boolean as <c>true</c>/<c>false</c>.
/// </summary>
public sealed record CustomPropertyDto(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("value")] string Value);

/// <summary>
/// Through the raw XML, like the other property parts. A property the editor cannot show (a vector,
/// a blob) is not in the model, and saving keeps it as it was.
/// </summary>
internal static class CustomProperties
{
    private static readonly XNamespace Custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    /// <summary>Word's identifier for user-defined properties: every one of them carries it.</summary>
    private const string UserDefined = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    /// <summary>Pids 0 and 1 are reserved by the property set format.</summary>
    private const int FirstPid = 2;

    private const string Text = "text";
    private const string Number = "number";
    private const string Date = "date";
    private const string Boolean = "boolean";

    private static readonly HashSet<string> TextTypes = new(StringComparer.Ordinal) { "lpwstr", "lpstr", "bstr" };

    private static readonly HashSet<string> IntegerTypes = new(StringComparer.Ordinal)
    {
        "i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint",
    };

    private static readonly HashSet<string> NumberTypes = new(StringComparer.Ordinal) { "r4", "r8", "decimal" };

    /// <summary>Null when the package has no custom part.</summary>
    public static List<CustomPropertyDto>? Read(OpenXmlPackage document)
    {
        var root = Load(PartOf(document))?.Root;
        return root?.Elements(Custom + "property").Select(Model).OfType<CustomPropertyDto>().ToList();
    }

    /// <summary>
    /// The list replaces the file's: a field with the same value keeps its element, a changed one gets
    /// a new value, and a new one goes at the end. Without any property left, the part goes.
    /// </summary>
    public static void Apply(OpenXmlPackage document, List<CustomPropertyDto>? wanted, HashSet<string> touched)
    {
        if (wanted is null) return;
        wanted = wanted.DistinctBy(property => property.Name, StringComparer.Ordinal).ToList();

        var part = PartOf(document);
        var xml = Load(part);
        var current = xml?.Root?.Elements(Custom + "property").ToList() ?? [];
        if (Same(current, wanted)) return;

        xml ??= New();
        var root = xml.Root!;
        Sync(root, current, wanted);
        touched.Add("_rels/.rels");

        if (!root.Elements(Custom + "property").Any())
        {
            if (part is not null) document.DeletePart(part);
            return;
        }

        part ??= document.AddNewPart<CustomFilePropertiesPart>();
        using (var stream = part.GetStream(FileMode.Create, FileAccess.Write)) XmlParts.Write(stream, xml);
        touched.Add(part.Uri.ToString().TrimStart('/'));
    }

    private static bool Same(List<XElement> current, List<CustomPropertyDto> wanted)
    {
        // In any order: the list on screen need not follow the file's.
        var shown = current.Select(Model).OfType<CustomPropertyDto>().ToList();
        return shown.Count == wanted.Count && shown.All(wanted.Contains);
    }

    private static void Sync(XElement root, List<XElement> current, List<CustomPropertyDto> wanted)
    {
        var byName = wanted.ToDictionary(property => property.Name, StringComparer.Ordinal);
        foreach (var element in current)
        {
            if (Model(element) is not { } shown) continue;
            if (!byName.TryGetValue(shown.Name, out var target)) element.Remove();
            else if (target != shown) SetValue(element, target);
        }

        var present = root.Elements(Custom + "property")
            .Select(element => (string?)element.Attribute("name"))
            .ToHashSet(StringComparer.Ordinal);
        var pid = NextPid(root);
        foreach (var property in wanted.Where(property => !present.Contains(property.Name)))
        {
            var element = new XElement(
                Custom + "property",
                new XAttribute("fmtid", UserDefined),
                new XAttribute("pid", (pid++).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("name", property.Name));
            SetValue(element, property);
            root.Add(element);
        }
    }

    private static int NextPid(XElement root) =>
        root.Elements(Custom + "property")
            .Select(element => int.TryParse((string?)element.Attribute("pid"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var pid) ? pid : 0)
            .DefaultIfEmpty(FirstPid - 1)
            .Max() + 1;

    /// <summary>Null for a type the editor does not show.</summary>
    private static CustomPropertyDto? Model(XElement element)
    {
        var name = (string?)element.Attribute("name");
        var value = element.Elements().FirstOrDefault(child => child.Name.Namespace == Vt);
        if (string.IsNullOrEmpty(name) || value is null) return null;

        var tag = value.Name.LocalName;
        if (TextTypes.Contains(tag)) return new CustomPropertyDto(name, Text, value.Value);
        if (IntegerTypes.Contains(tag) || NumberTypes.Contains(tag))
        {
            return double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
                ? new CustomPropertyDto(name, Number, value.Value.Trim())
                : null;
        }

        if (tag is "filetime" or "date") return new CustomPropertyDto(name, Date, value.Value.Trim());
        if (tag == "bool")
        {
            var flag = value.Value.Trim() is "true" or "1";
            return new CustomPropertyDto(name, Boolean, flag ? "true" : "false");
        }

        return null;
    }

    /// <summary>A number keeps the file's integer type when it still fits it.</summary>
    private static void SetValue(XElement element, CustomPropertyDto property)
    {
        var previous = element.Elements().FirstOrDefault(child => child.Name.Namespace == Vt)?.Name.LocalName;
        element.RemoveNodes();
        element.Add(new XElement(Vt + TagOf(property, previous), property.Value));
    }

    private static string TagOf(CustomPropertyDto property, string? previous) => property.Type switch
    {
        Number when IsInteger(property.Value) => previous is not null && IntegerTypes.Contains(previous) ? previous : "i4",
        Number => "r8",
        Date => "filetime",
        Boolean => "bool",
        _ => "lpwstr",
    };

    private static bool IsInteger(string value) =>
        int.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _);

    private static XDocument New() => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement(
            Custom + "Properties",
            new XAttribute("xmlns", Custom.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "vt", Vt.NamespaceName)));

    private static CustomFilePropertiesPart? PartOf(OpenXmlPackage document) =>
        document.GetPartsOfType<CustomFilePropertiesPart>().FirstOrDefault();

    private static XDocument? Load(OpenXmlPart? part)
    {
        if (part is null) return null;
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XmlParts.Read(stream);
        }
        catch (Exception problem) when (problem is IOException or InvalidOperationException)
        {
            return null;
        }
    }
}

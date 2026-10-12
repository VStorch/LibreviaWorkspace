using System.Globalization;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <c>docProps/core.xml</c> and part of <c>docProps/app.xml</c>. Each field is a patch: null leaves
/// the file as it is, and empty clears it. <c>TotalTime</c> is read only.
public sealed record PropertiesDto(
    [property: JsonPropertyName("title")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Title = null,
    [property: JsonPropertyName("subject")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Subject = null,
    [property: JsonPropertyName("creator")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Creator = null,
    [property: JsonPropertyName("keywords")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Keywords = null,
    [property: JsonPropertyName("category")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Category = null,
    [property: JsonPropertyName("description")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Description = null,
    [property: JsonPropertyName("lastModifiedBy")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? LastModifiedBy = null,
    [property: JsonPropertyName("revision")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Revision = null,
    [property: JsonPropertyName("created")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Created = null,
    [property: JsonPropertyName("modified")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Modified = null,
    [property: JsonPropertyName("company")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Company = null,
    [property: JsonPropertyName("manager")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? Manager = null,
    [property: JsonPropertyName("totalTime")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    int? TotalTime = null,
    /// See <see cref="Privacy.PrivacyFlag"/>. Only true is a patch: saving never clears the flag.
    [property: JsonPropertyName("removePersonalInformation")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    bool? RemovePersonalInformation = null,
    /// <c>docProps/custom.xml</c>: null leaves the part as it is; a list replaces it.
    [property: JsonPropertyName("custom")]
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    List<CustomPropertyDto>? Custom = null);

/// <summary>
/// Through the raw XML: the SDK's <c>PackageProperties</c> reserializes <c>core.xml</c> and drops
/// what it does not know. The same parts in a document and in a spreadsheet.
/// </summary>
internal static class DocumentProperties
{
    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTerms = "http://purl.org/dc/terms/";
    private static readonly XNamespace DcmiType = "http://purl.org/dc/dcmitype/";
    private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";
    private static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    private static readonly (XName Name, Func<PropertiesDto, string?> Get)[] CoreFields =
    [
        (Dc + "title", p => p.Title),
        (Dc + "subject", p => p.Subject),
        (Dc + "creator", p => p.Creator),
        (Cp + "keywords", p => p.Keywords),
        (Cp + "category", p => p.Category),
        (Dc + "description", p => p.Description),
        (Cp + "lastModifiedBy", p => p.LastModifiedBy),
        (Cp + "revision", p => p.Revision),
        (DcTerms + "created", p => p.Created),
        (DcTerms + "modified", p => p.Modified),
    ];

    private static readonly (XName Name, Func<PropertiesDto, string?> Get)[] AppFields =
    [
        (Ep + "Company", p => p.Company),
        (Ep + "Manager", p => p.Manager),
    ];

    /// <summary>Null when the package has neither part, or no field in them.</summary>
    public static PropertiesDto? Read(OpenXmlPackage document)
    {
        var core = Load(CorePart(document))?.Root;
        var app = Load(AppPart(document))?.Root;
        var anonymous = Privacy.PrivacyFlag.IsSet(Load(FlagPart(document)));
        var custom = CustomProperties.Read(document);
        if (core is null && app is null && !anonymous && custom is null) return null;

        // Empty means absent, not "clear": the new package's `<dc:creator/>` stays.
        string? Text(XElement? root, XName name) =>
            root?.Element(name)?.Value is { Length: > 0 } value ? value : null;

        int? totalTime = int.TryParse(
            Text(app, Ep + "TotalTime"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var minutes) &&
            minutes >= 0
            ? minutes
            : null;

        var read = new PropertiesDto(
            Title: Text(core, Dc + "title"),
            Subject: Text(core, Dc + "subject"),
            Creator: Text(core, Dc + "creator"),
            Keywords: Text(core, Cp + "keywords"),
            Category: Text(core, Cp + "category"),
            Description: Text(core, Dc + "description"),
            LastModifiedBy: Text(core, Cp + "lastModifiedBy"),
            Revision: Text(core, Cp + "revision"),
            Created: Text(core, DcTerms + "created"),
            Modified: Text(core, DcTerms + "modified"),
            Company: Text(app, Ep + "Company"),
            Manager: Text(app, Ep + "Manager"),
            TotalTime: totalTime,
            RemovePersonalInformation: anonymous ? true : null,
            Custom: custom is { Count: > 0 } ? custom : null);

        return read == new PropertiesDto() ? null : read;
    }

    /// <summary>
    /// Only the part that changes is rewritten, with what the editor does not know.
    /// </summary>
    public static void Apply(OpenXmlPackage document, PropertiesDto? model, HashSet<string> touched)
    {
        if (model is null) return;

        var corePart = CorePart(document);
        var core = Load(corePart);
        if (Patch(ref core, CoreFields, model, NewCore))
        {
            if (corePart is null)
            {
                corePart = document.AddNewPart<CoreFilePropertiesPart>();
                // The new relationship lives in `_rels/.rels`, which would come back without it.
                touched.Add("_rels/.rels");
            }

            Save(corePart, core!);
            touched.Add(corePart.Uri.ToString().TrimStart('/'));
        }

        var appPart = AppPart(document);
        var app = Load(appPart);
        if (Patch(ref app, AppFields, model, NewApp))
        {
            if (appPart is null)
            {
                appPart = document.AddNewPart<ExtendedFilePropertiesPart>();
                touched.Add("_rels/.rels");
            }

            Save(appPart, app!);
            touched.Add(appPart.Uri.ToString().TrimStart('/'));
        }

        if (model.RemovePersonalInformation == true) ApplyFlag(document, touched);
        CustomProperties.Apply(document, model.Custom, touched);
    }

    /// <summary>
    /// Through the typed DOM: saving tracked changes may already have loaded <c>settings.xml</c>,
    /// and the SDK would write that copy over a raw edit.
    /// </summary>
    private static void ApplyFlag(OpenXmlPackage document, HashSet<string> touched)
    {
        OpenXmlPart? changed = document switch
        {
            WordprocessingDocument { MainDocumentPart: { } main } => FlagWord(main),
            SpreadsheetDocument { WorkbookPart: { } workbook } => FlagWorkbook(workbook),
            _ => null,
        };
        if (changed is not null) touched.Add(changed.Uri.ToString().TrimStart('/'));
    }

    private static DocumentSettingsPart? FlagWord(MainDocumentPart main)
    {
        var part = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        var settings = part.Settings ??= new Wordprocessing.Settings();
        if (settings.GetFirstChild<Wordprocessing.RemovePersonalInformation>() is not null) return null;
        if (!settings.AddChild(new Wordprocessing.RemovePersonalInformation(), throwOnError: false)) return null;

        settings.Save();
        return part;
    }

    private static WorkbookPart? FlagWorkbook(WorkbookPart part)
    {
        if (part.Workbook is not { } workbook) return null;
        var properties = workbook.WorkbookProperties ??= new Spreadsheet.WorkbookProperties();
        if (properties.FilterPrivacy?.Value == true) return null;

        properties.FilterPrivacy = true;
        workbook.Save();
        return part;
    }

    /// <summary>Word's <c>settings.xml</c>, Excel's <c>workbook.xml</c>.</summary>
    private static OpenXmlPart? FlagPart(OpenXmlPackage document) => document switch
    {
        WordprocessingDocument word => word.MainDocumentPart?.DocumentSettingsPart,
        SpreadsheetDocument sheet => sheet.WorkbookPart,
        _ => null,
    };

    private static CoreFilePropertiesPart? CorePart(OpenXmlPackage document) =>
        document.GetPartsOfType<CoreFilePropertiesPart>().FirstOrDefault();

    private static ExtendedFilePropertiesPart? AppPart(OpenXmlPackage document) =>
        document.GetPartsOfType<ExtendedFilePropertiesPart>().FirstOrDefault();

    /// <returns>Whether any field changed.</returns>
    private static bool Patch(
        ref XDocument? xml,
        (XName Name, Func<PropertiesDto, string?> Get)[] fields,
        PropertiesDto model,
        Func<XDocument> create)
    {
        var changed = false;
        foreach (var (name, get) in fields)
        {
            var wanted = get(model);
            if (wanted is null) continue;

            var element = xml?.Root?.Element(name);
            if (wanted.Length == 0)
            {
                if (element is null) continue;
                element.Remove();
                changed = true;
                continue;
            }

            if (element is not null && element.Value == wanted && !element.HasElements) continue;

            xml ??= create();
            var root = xml.Root!;
            if (element is null)
            {
                element = new XElement(name);
                if (name.Namespace == DcTerms)
                {
                    // Without the type, Word reads `dcterms:created` as text, not as a date.
                    EnsurePrefix(root, "xsi", Xsi);
                    EnsurePrefix(root, "dcterms", DcTerms);
                    element.SetAttributeValue(Xsi + "type", "dcterms:W3CDTF");
                }

                root.Add(element);
            }

            element.RemoveNodes();
            element.Value = wanted;
            changed = true;
        }

        return changed;
    }

    private static void EnsurePrefix(XElement root, string prefix, XNamespace ns)
    {
        if (root.GetPrefixOfNamespace(ns) is not null) return;
        root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
    }

    private static XDocument NewCore() => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement(
            Cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", Cp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc", Dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcterms", DcTerms.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcmitype", DcmiType.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsi", Xsi.NamespaceName)));

    private static XDocument NewApp() => new(
        new XDeclaration("1.0", "UTF-8", "yes"),
        new XElement(
            Ep + "Properties",
            new XAttribute("xmlns", Ep.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "vt", Vt.NamespaceName)));

    /// <summary>
    /// Null when the part is missing or malformed: properties are accessory, and do not refuse the
    /// document.
    /// </summary>
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

    private static void Save(OpenXmlPart part, XDocument xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        XmlParts.Write(stream, xml);
    }
}

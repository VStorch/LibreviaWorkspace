using System.Globalization;
using System.Text;
using System.Text.Json.Serialization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Librevia.Format.Docx;

/// <summary>
/// <c>docProps/core.xml</c> e parte de <c>docProps/app.xml</c>. Cada campo é um
/// remendo: nulo deixa o arquivo como está, e vazio apaga. <c>TotalTime</c> só é lido.
/// </summary>
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
    int? TotalTime = null);

/// <summary>
/// Pelo XML cru: o <c>PackageProperties</c> do SDK reserializa <c>core.xml</c> e
/// descarta o que não conhece.
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

    /// <summary>Nulo quando o pacote não tem nenhuma das duas partes, ou nenhum campo nelas.</summary>
    public static PropertiesDto? Read(WordprocessingDocument document)
    {
        var core = Load(document.CoreFilePropertiesPart)?.Root;
        var app = Load(document.ExtendedFilePropertiesPart)?.Root;
        if (core is null && app is null) return null;

        // Vazio é ausente, e não "apague": o `<dc:creator/>` do pacote novo fica.
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
            TotalTime: totalTime);

        return read == new PropertiesDto() ? null : read;
    }

    /// <summary>Só a parte que muda é regravada, com o que o editor não conhece.</summary>
    public static void Apply(WordprocessingDocument document, PropertiesDto? model, HashSet<string> touched)
    {
        if (model is null) return;

        var corePart = document.CoreFilePropertiesPart;
        var core = Load(corePart);
        if (Patch(ref core, CoreFields, model, NewCore))
        {
            if (corePart is null)
            {
                corePart = document.AddCoreFilePropertiesPart();
                // A relação nova mora em `_rels/.rels`, que voltaria sem ela.
                touched.Add("_rels/.rels");
            }

            Save(corePart, core!);
            touched.Add(corePart.Uri.ToString().TrimStart('/'));
        }

        var appPart = document.ExtendedFilePropertiesPart;
        var app = Load(appPart);
        if (Patch(ref app, AppFields, model, NewApp))
        {
            if (appPart is null)
            {
                appPart = document.AddExtendedFilePropertiesPart();
                touched.Add("_rels/.rels");
            }

            Save(appPart, app!);
            touched.Add(appPart.Uri.ToString().TrimStart('/'));
        }
    }

    /// <returns>Se algum campo mudou.</returns>
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
                    // Sem o tipo, o Word lê `dcterms:created` como texto, e não como data.
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
    /// Nulo quando a parte falta ou está malformada: as propriedades são acessórias, e
    /// não recusam o documento.
    /// </summary>
    private static XDocument? Load(OpenXmlPart? part)
    {
        if (part is null) return null;
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = XmlReader.Create(
                stream,
                new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (Exception problem) when (problem is XmlException or IOException or InvalidOperationException)
        {
            return null;
        }
    }

    private static void Save(OpenXmlPart part, XDocument xml)
    {
        // O XmlWriter escreveria `utf-8`; o Word e o LibreOffice escrevem `UTF-8`.
        var standalone = xml.Declaration?.Standalone is { Length: > 0 } value ? value : "yes";
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        var declaration = Encoding.UTF8.GetBytes(
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"{standalone}\"?>\r\n");
        stream.Write(declaration, 0, declaration.Length);
        using var writer = XmlWriter.Create(
            stream,
            new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = true });
        xml.Root!.Save(writer);
    }
}

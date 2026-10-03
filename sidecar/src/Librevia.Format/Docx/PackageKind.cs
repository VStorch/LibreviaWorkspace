using System.IO.Compression;
using System.Xml.Linq;

namespace Librevia.Format.Docx;

/// <summary>
/// Documento ou modelo: o tipo do pacote, dito pelo tipo de conteúdo da parte principal.
/// </summary>
/// <remarks>
/// Um `.dotx` é um `.docx` com outro rótulo em `[Content_Types].xml` — o mesmo
/// `word/document.xml`, os mesmos estilos, cabeçalhos e tema. O documento
/// criado a partir de um modelo parte dos bytes do modelo, e é a gravação que
/// troca o rótulo para o do destino: gravado como `.docx` com o rótulo de
/// modelo, o Word recusa o arquivo.
///
/// O `.dotm` leva macros (VBA). Elas não viajam: o aplicativo não as executa nem as
/// edita, e um `.docx` com `vbaProject.bin` é inválido. Saem o projeto, os dados
/// dele e as personalizações de teclado e barra que só existem em pacote com macro
/// — e a perda é declarada, nunca calada.
///
/// Trabalha nos bytes do zip, depois da gravação cirúrgica: as partes que não são
/// do rótulo nem das macros saem como entraram.
/// </remarks>
public static class PackageKind
{
    public const string DocumentMain =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    public const string TemplateMain =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
    public const string MacroDocumentMain = "application/vnd.ms-word.document.macroEnabled.main+xml";
    public const string MacroTemplateMain = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";

    /// <summary>A frase da perda das macros, na abertura e na gravação.</summary>
    public const string Macros = "macros (VBA) do modelo";

    private static readonly XNamespace ContentTypes =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Relationships =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string OfficeDocument =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    /// <summary>As relações que só existem em pacote com macro.</summary>
    private static readonly HashSet<string> MacroRelationships = new(StringComparer.Ordinal)
    {
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
        "http://schemas.microsoft.com/office/2006/relationships/wordVbaData",
        "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations",
        "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars",
    };

    /// <summary>O pacote traz macros — o `.dotm` (ou um `.docm` renomeado).</summary>
    public static bool HasMacros(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var main = MainPartName(archive);
        return main is not null && MainContentType(archive, main) is MacroDocumentMain or MacroTemplateMain;
    }

    /// <summary>
    /// Devolve o pacote com o rótulo de documento ou de modelo, sem as macros.
    /// </summary>
    /// <returns>Os mesmos bytes quando já estava certo — nada a mexer.</returns>
    public static byte[] Retarget(byte[] bytes, bool template, Inventory inventory)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var main = MainPartName(archive);
        if (main is null) return bytes;

        var current = MainContentType(archive, main);
        var wanted = template ? TemplateMain : DocumentMain;
        if (string.Equals(current, wanted, StringComparison.Ordinal)) return bytes;

        var types = ReadXml(archive, "[Content_Types].xml")!;
        var dropped = new HashSet<string>(StringComparer.Ordinal);
        var rewritten = new Dictionary<string, XDocument>(StringComparer.Ordinal);

        if (current is MacroDocumentMain or MacroTemplateMain)
        {
            // As relações de macro saem de toda parte que as tenha (o projeto aponta
            // os dados dele), e as partes apontadas saem do pacote.
            foreach (var entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".rels", StringComparison.Ordinal)))
            {
                var rels = ReadXml(archive, entry.FullName)!;
                var macro = rels.Root!.Elements(Relationships + "Relationship")
                    .Where(rel => MacroRelationships.Contains((string?)rel.Attribute("Type") ?? string.Empty))
                    .ToList();
                if (macro.Count == 0) continue;

                var owner = OwnerOf(entry.FullName);
                foreach (var rel in macro)
                {
                    if ((string?)rel.Attribute("TargetMode") == "External") continue;
                    dropped.Add(Resolve(owner, (string?)rel.Attribute("Target") ?? string.Empty));
                }

                macro.ForEach(rel => rel.Remove());
                rewritten[entry.FullName] = rels;
            }

            // As relações das partes que saíram saem junto com elas.
            foreach (var part in dropped.ToList()) dropped.Add(RelsOf(part));
            inventory.NoteLoss(Macros);
        }

        foreach (var name in rewritten.Keys.Where(dropped.Contains).ToList()) rewritten.Remove(name);

        // O rótulo da parte principal, e as declarações das partes que saíram.
        var labeled = false;
        foreach (var over in types.Root!.Elements(ContentTypes + "Override").ToList())
        {
            var part = ((string?)over.Attribute("PartName") ?? string.Empty).TrimStart('/');
            if (string.Equals(part, main, StringComparison.OrdinalIgnoreCase))
            {
                over.SetAttributeValue("ContentType", wanted);
                labeled = true;
            }
            else if (dropped.Contains(part))
            {
                over.Remove();
            }
        }

        // Rotulada pelo padrão da extensão (o SDK faz isso no pacote de uma parte
        // só): ganha uma declaração própria, que vale mais que o padrão.
        if (!labeled)
        {
            types.Root.Add(new XElement(
                ContentTypes + "Override",
                new XAttribute("PartName", "/" + main),
                new XAttribute("ContentType", wanted)));
        }

        // O padrão do `.bin` é o do projeto VBA: só sai quando não sobra `.bin` nenhum.
        var remaining = archive.Entries.Select(entry => entry.FullName).Where(name => !dropped.Contains(name)).ToList();
        foreach (var fallback in types.Root.Elements(ContentTypes + "Default").ToList())
        {
            var extension = (string?)fallback.Attribute("Extension") ?? string.Empty;
            // O padrão que o SDK põe no `.xml` do pacote de uma parte só é o rótulo
            // da principal: troca junto, para que nenhum rótulo antigo sobre.
            if ((string?)fallback.Attribute("ContentType") is DocumentMain or TemplateMain or MacroDocumentMain or MacroTemplateMain)
            {
                fallback.SetAttributeValue("ContentType", wanted);
            }
            else if ((string?)fallback.Attribute("ContentType") == "application/vnd.ms-office.vbaProject" &&
                !remaining.Any(name => name.EndsWith("." + extension, StringComparison.OrdinalIgnoreCase)))
            {
                fallback.Remove();
            }
        }

        rewritten["[Content_Types].xml"] = types;

        using var result = new MemoryStream();
        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in archive.Entries)
            {
                if (dropped.Contains(entry.FullName)) continue;

                var target = output.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                target.LastWriteTime = entry.LastWriteTime;
                using var stream = target.Open();
                if (rewritten.TryGetValue(entry.FullName, out var xml))
                {
                    xml.Save(stream, SaveOptions.DisableFormatting);
                }
                else
                {
                    using var source = entry.Open();
                    source.CopyTo(stream);
                }
            }
        }

        return result.ToArray();
    }

    /// <summary>A parte principal, pela relação `officeDocument` do pacote.</summary>
    private static string? MainPartName(ZipArchive archive)
    {
        var rels = ReadXml(archive, "_rels/.rels");
        var target = rels?.Root?.Elements(Relationships + "Relationship")
            .FirstOrDefault(rel => (string?)rel.Attribute("Type") == OfficeDocument)
            ?.Attribute("Target")?.Value;
        return target is null ? null : Resolve(string.Empty, target);
    }

    private static string? MainContentType(ZipArchive archive, string main)
    {
        var types = ReadXml(archive, "[Content_Types].xml");
        var extension = main[(main.LastIndexOf('.') + 1)..];
        return types?.Root?.Elements(ContentTypes + "Override")
                   .FirstOrDefault(over => string.Equals(
                       ((string?)over.Attribute("PartName") ?? string.Empty).TrimStart('/'), main,
                       StringComparison.OrdinalIgnoreCase))
                   ?.Attribute("ContentType")?.Value
               ?? types?.Root?.Elements(ContentTypes + "Default")
                   .FirstOrDefault(fallback => string.Equals(
                       (string?)fallback.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase))
                   ?.Attribute("ContentType")?.Value;
    }

    private static XDocument? ReadXml(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        if (entry is null) return null;
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>A pasta da parte dona de um `.rels` (`word/_rels/document.xml.rels` → `word`).</summary>
    private static string OwnerOf(string rels)
    {
        var folder = rels[..Math.Max(0, rels.LastIndexOf("_rels/", StringComparison.Ordinal))].TrimEnd('/');
        return folder;
    }

    private static string RelsOf(string part)
    {
        var slash = part.LastIndexOf('/');
        return slash < 0 ? $"_rels/{part}.rels" : $"{part[..slash]}/_rels/{part[(slash + 1)..]}.rels";
    }

    /// <summary>O alvo de uma relação, relativo à pasta da dona, como nome de entrada do zip.</summary>
    private static string Resolve(string folder, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = (folder.Length == 0 ? [] : folder.Split('/')).ToList();
        foreach (var segment in target.Split('/'))
        {
            if (segment == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
            else if (segment != "." && segment.Length > 0) parts.Add(segment);
        }

        return string.Join('/', parts);
    }
}

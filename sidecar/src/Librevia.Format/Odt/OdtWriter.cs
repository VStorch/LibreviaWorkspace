using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Librevia.Format.Docx;

namespace Librevia.Format.Odt;

/// <summary>
/// <c>odt.save</c>, surgical as <see cref="DocxWriter"/>: the original is read again with the same
/// reader, and only what changed is written. A block nobody touched goes back as the exact text it had
/// in <c>content.xml</c>; <c>styles.xml</c> changes only for a named style or the page,
/// <c>meta.xml</c> only for the properties, the manifest only for a new image or a change between
/// document and template (<see cref="OdtMediaType"/>). Every other entry
/// comes back with the same content, and <c>mimetype</c> first and stored, as ODF asks.
/// </summary>
public static class OdtWriter
{
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);

    public static (byte[] Bytes, SaveResult Result) Write(byte[] original, DocumentModelDto model)
    {
        using var stream = new MemoryStream(original, writable: false);
        using var archive = OdtReader.OpenArchive(stream);
        var package = OdtPackage.Open(archive);
        var source = OdtSource.Read(package);

        var inventory = new Inventory();
        var pictures = new OdtPictures(package);
        var taken = OdtAutomaticStyles.NamesIn(package.Content, package.Styles).ToList();
        var context = new OdtWriteContext(source.Styles, new OdtAutomaticStyles(taken), pictures, inventory)
        {
            Plan = new OdtDocumentPlan(model.Doc, model.Comments ?? source.Comments, model.Notes ?? source.Notes, source.CommentNames, package.Content),
            Formulas = new OdtFormulas(path => package.Has(path) || package.MediaTypes.ContainsKey(path + "/")),
            Changes = new OdtChangeWriter(source.Changes, package.Content.Root!, model.TrackChanges),
        };
        context.TakeNames(package.Content.Root!);

        var replaced = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var masters = OdtSectionMasters.Assign(source, [.. model.Sections ?? [], model.Page], MastersOf(package.Styles));
        var (content, preserved, rewritten) = Content(package, source, context, model, masters);
        var styles = OdtStylesPart.Write(source, model, taken, masters.Changes);
        Replace(replaced, package, OdtPackage.ContentPath, content);
        Replace(replaced, package, OdtPackage.StylesPath, styles);
        if (OdtMetaPart.Write(package.Meta, source.Properties, model.Properties) is { } meta) replaced[OdtPackage.MetaPath] = meta;
        var objects = OdtEmbeddedObjects.Find(package, content, styles is null ? package.Styles : XmlParts.Read(Utf8.GetBytes(styles)));
        var addMeta = package.Meta is null && replaced.ContainsKey(OdtPackage.MetaPath);
        var mediaType = OdtMediaType.Wanted(package, model.Template);
        if (mediaType is not null && package.Has(OdtPackage.MimeTypePath)) replaced[OdtPackage.MimeTypePath] = OdtMediaType.Bytes(mediaType);
        var changes = new ManifestChanges(pictures, context.Formulas, objects, addMeta, mediaType);
        Replace(replaced, package, OdtPackage.ManifestPath, Manifest(package, changes));

        return (Zip(package, replaced, pictures, context.Formulas, objects), new SaveResult(inventory, preserved, rewritten));
    }

    /// <summary>Only when the text differs: an unchanged part keeps its bytes.</summary>
    private static void Replace(Dictionary<string, byte[]> replaced, OdtPackage package, string path, string? text)
    {
        if (text is null) return;
        var bytes = Utf8.GetBytes(text);
        if (package.ReadBytes(path) is { } before && before.AsSpan().SequenceEqual(bytes)) return;
        replaced[path] = bytes;
    }

    internal static string TextOf(OdtPackage package, string path) =>
        Encoding.UTF8.GetString(package.ReadBytes(path) ?? throw new OdtException(OdtException.WithoutContent));

    private static (string Text, int Preserved, int Rewritten) Content(
        OdtPackage package, OdtSource source, OdtWriteContext context, DocumentModelDto model, OdtSectionMasters masters)
    {
        var xml = TextOf(package, OdtPackage.ContentPath);
        var root = package.Content.Root!;
        var spans = OdtXmlSpans.Root(xml);
        var xmlOut = new OdtXmlOut(root);
        var textSpan = OdtXmlSpans.Find(spans, source.Text);
        var sections = new OdtSectionPlan(model, masters, context, xmlOut);
        var body = new OdtBodyWriter(source, context, xmlOut, xml, textSpan, sections);
        var inner = body.Write(model.Doc);

        var edits = new TextEdits(xml);
        edits.ReplaceContent(textSpan, inner);
        AppendStyles(edits, spans, root, xmlOut, context.Automatic.Created, root.Element(Ns.Office + "body"));
        return (edits.Apply(), body.Preserved, body.Rewritten);
    }

    private static HashSet<string> MastersOf(XDocument? styles) =>
        (styles?.Root?.Element(Ns.Office + "master-styles")?.Elements(Ns.Style + "master-page") ?? [])
        .Select(master => master.Attr(Ns.Style + "name")).OfType<string>().ToHashSet(StringComparer.Ordinal);

    /// <summary>At the end of <c>office:automatic-styles</c>, created before <paramref name="before"/> when missing.</summary>
    internal static void AppendStyles(
        TextEdits edits, XmlSpan spans, XElement root, OdtXmlOut xmlOut, IReadOnlyList<XElement> styles, XElement? before)
    {
        if (styles.Count == 0) return;
        if (root.Element(Ns.Office + "automatic-styles") is { } automatic)
        {
            edits.AppendContent(OdtXmlSpans.Find(spans, automatic), xmlOut.Write(styles));
            return;
        }

        var created = xmlOut.Write(new XElement(Ns.Office + "automatic-styles", styles));
        if (before is not null) edits.Insert(OdtXmlSpans.Find(spans, before).Start, created);
        else edits.AppendContent(spans, created);
    }

    /// <summary>What a save adds to or removes from the manifest; <c>MediaType</c> is the root's new type.</summary>
    private sealed record ManifestChanges(
        OdtPictures Pictures, OdtFormulas Formulas, OdtEmbeddedObjects Objects, bool AddMeta, string? MediaType)
    {
        public bool None => Pictures.Added.Count == 0 && Formulas.Added.Count == 0 && !AddMeta && Objects.IsEmpty && MediaType is null;
    }

    private static string? Manifest(OdtPackage package, ManifestChanges changes)
    {
        if (changes.None || package.Manifest?.Root is not { } root) return null;
        var (pictures, formulas, objects, addMeta, mediaType) = changes;
        var entries = pictures.ManifestEntries().Concat(formulas.ManifestEntries()).ToList();
        if (addMeta)
        {
            entries.Add(new XElement(
                Ns.Manifest + "file-entry",
                new XAttribute(Ns.Manifest + "full-path", OdtPackage.MetaPath),
                new XAttribute(Ns.Manifest + "media-type", "text/xml")));
        }

        var xml = TextOf(package, OdtPackage.ManifestPath);
        var edits = new TextEdits(xml);
        var spans = OdtXmlSpans.Root(xml);
        foreach (var entry in root.Elements(Ns.Manifest + "file-entry"))
        {
            if (entry.Attr(Ns.Manifest + "full-path") is not { } path) continue;
            var span = OdtXmlSpans.Find(spans, entry);
            if (objects.Drops(path)) edits.Replace(span, string.Empty);
            else if (path == OdtMediaType.RootPath && mediaType is not null)
            {
                edits.Replace(span, OdtMediaType.Retyped(xml[span.Start..span.End], entry.Attr(Ns.Manifest + "media-type"), mediaType));
            }
        }

        if (entries.Count > 0) edits.AppendContent(spans, new OdtXmlOut(root).Write(entries));
        return edits.Apply();
    }

    private static byte[] Zip(
        OdtPackage package, Dictionary<string, byte[]> replaced, OdtPictures pictures, OdtFormulas formulas, OdtEmbeddedObjects objects)
    {
        var added = pictures.Added.Select(picture => new OdtAddedEntry(picture.Path, picture.Bytes, CompressionLevel.NoCompression))
            .Concat(formulas.Added.Select(formula => new OdtAddedEntry(OdtFormulas.ContentPathOf(formula.Folder), formula.Content, CompressionLevel.Optimal)))
            .ToList();
        if (!package.Has(OdtPackage.MetaPath) && replaced.TryGetValue(OdtPackage.MetaPath, out var meta))
        {
            added.Add(new OdtAddedEntry(OdtPackage.MetaPath, meta, CompressionLevel.Optimal));
        }

        return OdtZip.Write(
            package,
            path => objects.Drops(path) ? null : replaced.GetValueOrDefault(path) ?? package.ReadBytes(path) ?? [],
            added);
    }
}

/// <summary>Replacements in a part's text by position, applied from the end so positions hold.</summary>
internal sealed class TextEdits(string text)
{
    private readonly List<(int Start, int End, string Replacement)> _edits = [];

    public void Replace(int start, int end, string replacement) => _edits.Add((start, end, replacement));

    public void Replace(XmlSpan span, string replacement) => Replace(span.Start, span.End, replacement);

    public void Insert(int at, string piece) => Replace(at, at, piece);

    /// <summary>Between the tags; a self-closing element is opened and closed around it.</summary>
    public void ReplaceContent(XmlSpan span, string inner)
    {
        if (!span.SelfClosing)
        {
            Replace(span.ContentStart, span.ContentEnd, inner);
            return;
        }

        if (inner.Length == 0) return;
        var tag = text[span.Start..span.End];
        var name = tag[1..].Split([' ', '\t', '\r', '\n', '/', '>'], 2)[0];
        Replace(span, tag[..^2].TrimEnd() + ">" + inner + "</" + name + ">");
    }

    public void AppendContent(XmlSpan span, string piece)
    {
        if (span.SelfClosing) ReplaceContent(span, piece);
        else Insert(span.ContentEnd, piece);
    }

    public string Apply()
    {
        var result = new StringBuilder(text);
        var limit = text.Length;
        foreach (var (start, end, replacement) in _edits.OrderByDescending(edit => edit.Start).ThenByDescending(edit => edit.End))
        {
            if (end > limit) throw new InvalidOperationException($"The edit at {start}..{end} overlaps the one at {limit}.");
            result.Remove(start, end - start).Insert(start, replacement);
            limit = start;
        }

        return result.ToString();
    }
}

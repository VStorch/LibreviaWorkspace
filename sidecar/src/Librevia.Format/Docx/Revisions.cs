using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A text revision is a range mark, with author, date and <c>w:id</c>; the date travels as the file
/// wrote it, without a <c>DateTime</c>, which would give it a time zone or seconds. The paragraph
/// mark's and the row's are block attributes (<c>markRevision</c>, <c>rowRevision</c>).
/// </summary>
public static class Revisions
{
    public const string Insertion = "insertion";
    public const string Deletion = "deletion";

    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>As in the file, or null.</summary>
    public static string? AttributeOf(OpenXmlElement element, string localName) =>
        element.GetAttributes()
            .FirstOrDefault(attribute => attribute.LocalName == localName && attribute.NamespaceUri == W)
            .Value;

    /// <summary>Through the SDK's typed properties, in schema order; the date as text.</summary>
    internal static void Stamp(OpenXmlElement element, string? id, string? author, string? date)
    {
        var when = date is null ? null : new DateTimeValue { InnerText = date };
        switch (element)
        {
            case InsertedRun run: if (id is not null) run.Id = id; if (author is not null) run.Author = author; if (when is not null) run.Date = when; break;
            case DeletedRun run: if (id is not null) run.Id = id; if (author is not null) run.Author = author; if (when is not null) run.Date = when; break;
            case Inserted mark: if (id is not null) mark.Id = id; if (author is not null) mark.Author = author; if (when is not null) mark.Date = when; break;
            case Deleted mark: if (id is not null) mark.Id = id; if (author is not null) mark.Author = author; if (when is not null) mark.Date = when; break;
            case MoveFromRun run: if (id is not null) run.Id = id; break;
            case MoveToRun run: if (id is not null) run.Id = id; break;
            case MoveFrom mark: if (id is not null) mark.Id = id; break;
            case MoveTo mark: if (id is not null) mark.Id = id; break;
        }
    }

    public static Mark MarkOf(OpenXmlElement revision, string? moveName)
    {
        var (type, move) = revision switch
        {
            DeletedRun => (Deletion, (string?)null),
            MoveFromRun => (Deletion, "from"),
            MoveToRun => (Insertion, "to"),
            _ => (Insertion, null),
        };

        var attrs = new Dictionary<string, JsonNode?>();
        if (AttributeOf(revision, "author") is { } author) attrs["author"] = author;
        if (AttributeOf(revision, "date") is { } date) attrs["date"] = date;
        if (AttributeOf(revision, "id") is { } rid) attrs["rid"] = rid;
        if (move is not null)
        {
            attrs["move"] = move;
            if (moveName is not null) attrs["moveName"] = moveName;
        }

        return new Mark { Type = type, Attrs = attrs.Count > 0 ? attrs : null };
    }

    /// <summary>A move becomes a deletion or an insertion.</summary>
    public static JsonObject? BlockRevisionOf(OpenXmlElement? properties)
    {
        var revision = properties?.ChildElements.FirstOrDefault(IsBlockRevision);
        if (revision is null) return null;

        var result = new JsonObject { ["kind"] = KindOf(revision) };
        if (AttributeOf(revision, "author") is { } author) result["author"] = author;
        if (AttributeOf(revision, "date") is { } date) result["date"] = date;
        if (AttributeOf(revision, "id") is { } rid) result["rid"] = rid;
        return result;
    }

    private static bool IsBlockRevision(OpenXmlElement element) =>
        element.NamespaceUri == W && element.LocalName is "ins" or "del" or "moveFrom" or "moveTo";

    private static string KindOf(OpenXmlElement revision) =>
        revision.LocalName is "del" or "moveFrom" ? "del" : "ins";

    /// <summary>
    /// The file's stays when it is the same (kind and id), with what the model does not carry.
    /// <paramref name="insert"/> places the element at its schema slot, which differs between
    /// <c>w:rPr</c> and <c>w:trPr</c>.
    /// </summary>
    public static void ApplyBlock(OpenXmlElement properties, JsonNode? wanted, Action<OpenXmlElement> insert)
    {
        var existing = properties.ChildElements.Where(IsBlockRevision).ToList();
        var kind = (wanted as JsonObject)?["kind"]?.GetValue<string>();
        var rid = (wanted as JsonObject)?["rid"]?.GetValue<string>();

        if (kind is not null && existing.Count == 1 && KindOf(existing[0]) == kind &&
            string.Equals(AttributeOf(existing[0], "id"), rid, StringComparison.Ordinal))
        {
            return;
        }

        foreach (var element in existing) element.Remove();
        if (kind is null || wanted is not JsonObject revision) return;

        OpenXmlElement created = kind == "del" ? new Deleted() : new Inserted();
        Stamp(created, rid, revision["author"]?.GetValue<string>(), revision["date"]?.GetValue<string>());
        insert(created);
    }

    public static bool HasRevision(Node node) =>
        node.Marks?.Any(mark => mark.Type is Insertion or Deletion) == true;

    /// <summary>
    /// The link stays outside: <c>w:hyperlink</c> may contain <c>w:ins</c>, not the other way
    /// around. Equal neighbours merge in <see cref="MergeNeighbours"/>.
    /// </summary>
    public static IEnumerable<OpenXmlElement> Wrap(List<OpenXmlElement> elements, List<Mark>? marks)
    {
        var inserted = marks?.FirstOrDefault(mark => mark.Type == Insertion);
        var deleted = marks?.FirstOrDefault(mark => mark.Type == Deletion);
        if (inserted is null && deleted is null) return elements;

        var result = new List<OpenXmlElement>();
        var pending = new List<OpenXmlElement>();

        void Flush()
        {
            if (pending.Count == 0) return;
            result.Add(Wrapped(pending, inserted, deleted));
            pending = [];
        }

        foreach (var element in elements)
        {
            if (element is Hyperlink link)
            {
                Flush();
                var inner = link.ChildElements.ToList();
                foreach (var child in inner) child.Remove();
                link.AppendChild(Wrapped(inner, inserted, deleted));
                result.Add(link);
                continue;
            }

            pending.Add(element);
        }

        Flush();
        return result;
    }

    private static OpenXmlElement Wrapped(List<OpenXmlElement> elements, Mark? inserted, Mark? deleted)
    {
        OpenXmlCompositeElement? outer = null;
        if (deleted is not null)
        {
            var del = WrapperOf(new DeletedRun(), deleted);
            foreach (var element in elements)
            {
                AsDeleted(element);
                del.AppendChild(element);
            }

            outer = del;
        }

        if (inserted is not null)
        {
            var ins = WrapperOf(new InsertedRun(), inserted);
            if (outer is not null) ins.AppendChild(outer);
            else foreach (var element in elements) ins.AppendChild(element);
            outer = ins;
        }

        return outer!;
    }

    private static OpenXmlCompositeElement WrapperOf(OpenXmlCompositeElement wrapper, Mark mark)
    {
        Stamp(wrapper, Attr.MarkString(mark, "rid"), Attr.MarkString(mark, "author"), Attr.MarkString(mark, "date"));
        return wrapper;
    }

    private static void AsDeleted(OpenXmlElement element)
    {
        foreach (var text in element.Descendants<Text>().ToList())
        {
            text.InsertAfterSelf(new DeletedText(text.Text) { Space = SpaceProcessingModeValues.Preserve });
            text.Remove();
        }

        foreach (var code in element.Descendants<FieldCode>().ToList())
        {
            code.InsertAfterSelf(new DeletedFieldCode(code.Text) { Space = SpaceProcessingModeValues.Preserve });
            code.Remove();
        }
    }

    /// <summary>
    /// Same id, author and date, as Word writes; descends into links and nested revisions.
    /// </summary>
    public static void MergeNeighbours(OpenXmlElement parent)
    {
        OpenXmlElement? previous = null;
        foreach (var child in parent.ChildElements.ToList())
        {
            if (child is InsertedRun or DeletedRun && previous is not null && Same(previous, child))
            {
                foreach (var inner in child.ChildElements.ToList())
                {
                    inner.Remove();
                    previous.AppendChild(inner);
                }

                child.Remove();
                continue;
            }

            previous = child;
        }

        foreach (var child in parent.ChildElements)
        {
            if (child is InsertedRun or DeletedRun or Hyperlink) MergeNeighbours(child);
        }
    }

    private static bool Same(OpenXmlElement a, OpenXmlElement b) =>
        a.GetType() == b.GetType() &&
        AttributeOf(a, "id") == AttributeOf(b, "id") &&
        AttributeOf(a, "author") == AttributeOf(b, "author") &&
        AttributeOf(a, "date") == AttributeOf(b, "date");

    /// <summary>
    /// The file's id stays while unique; a repeated one (split paragraph) or a missing one gets one
    /// above the package's highest <c>w:id</c>, which revisions, comments and bookmarks share.
    /// Notes enter the same count.
    /// </summary>
    public static void MakeIdsUnique(Body body, MainDocumentPart part)
    {
        IEnumerable<OpenXmlElement?> roots = [body, part.FootnotesPart?.Footnotes, part.EndnotesPart?.Endnotes];
        var revisions = roots.OfType<OpenXmlElement>().SelectMany(root => root.Descendants().Where(IsRevision)).ToList();
        if (revisions.Count == 0) return;

        var used = new HashSet<string>(StringComparer.Ordinal);
        int? next = null;
        foreach (var revision in revisions)
        {
            var id = AttributeOf(revision, "id");
            if (id is not null && used.Add(id)) continue;

            next ??= HighestId(body, part) + 1;
            var fresh = next.Value.ToString(CultureInfo.InvariantCulture);
            next++;
            used.Add(fresh);
            Stamp(revision, fresh, null, null);
        }
    }

    private static bool IsRevision(OpenXmlElement element) =>
        element.NamespaceUri == W && element.LocalName is "ins" or "del" or "moveFrom" or "moveTo";

    private static int HighestId(Body body, MainDocumentPart part)
    {
        IEnumerable<OpenXmlElement?> roots =
        [
            body,
            .. part.HeaderParts.Select(header => header.Header),
            .. part.FooterParts.Select(footer => footer.Footer),
            part.FootnotesPart?.Footnotes,
            part.EndnotesPart?.Endnotes,
            part.WordprocessingCommentsPart?.Comments,
        ];

        var highest = 0;
        foreach (var root in roots.OfType<OpenXmlElement>())
        {
            foreach (var element in root.Descendants())
            {
                if (AttributeOf(element, "id") is { } id &&
                    int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) &&
                    number > highest)
                {
                    highest = number;
                }
            }
        }

        return highest;
    }

    public static bool TrackingOf(MainDocumentPart part) =>
        part.DocumentSettingsPart?.Settings?.GetFirstChild<TrackRevisions>() is { } track &&
        !(track.Val is { } value && !value.Value);

    /// <summary>Only when it differs from the file.</summary>
    public static void ApplyTracking(MainDocumentPart part, bool? wanted, HashSet<string> touched, Inventory inventory)
    {
        if (wanted is not { } track || track == TrackingOf(part)) return;

        var settingsPart = part.DocumentSettingsPart ?? part.AddNewPart<DocumentSettingsPart>();
        var settings = settingsPart.Settings ??= new Settings();
        settings.RemoveAllChildren<TrackRevisions>();
        if (track && !settings.AddChild(new TrackRevisions(), throwOnError: false))
        {
            inventory.NoteLoss("\"Controlar alterações\" (o arquivo não aceitou o interruptor)");
            return;
        }

        settings.Save();
        touched.Add(settingsPart.Uri.ToString().TrimStart('/'));
    }
}

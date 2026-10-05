using System.IO.Compression;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Librevia.Format.Docx;

public sealed record SaveResult(
    [property: JsonPropertyName("inventory")] Inventory Inventory,
    [property: JsonPropertyName("preservedBlocks")] int PreservedBlocks,
    [property: JsonPropertyName("rewrittenBlocks")] int RewrittenBlocks);

/// <summary>
/// Surgical save: fidelity comes from **not touching** what was not edited. Only
/// <c>word/document.xml</c> is rewritten; other parts only change if the user changed them.
/// </summary>
public static class DocxWriter
{
    public static (byte[] Bytes, SaveResult Result) Write(byte[] original, DocumentModelDto model)
    {
        var inventory = new Inventory();

        // A writable copy: `original` is the bytes main kept on open.
        using var buffer = new MemoryStream();
        buffer.Write(original, 0, original.Length);
        buffer.Position = 0;

        using var document = OpenEditable(buffer);
        var part = document.MainDocumentPart
                   ?? throw new DocxException("O arquivo original não contém um documento do Word.");
        var body = part.Document?.Body
                   ?? throw new DocxException("O arquivo original está vazio ou danificado.");

        // Reindexes the same bytes with the same reader that produced the model: ids come out
        // equal, and a flattened draft is compared with a flattened reading.
        var reader = new BodyReader(
            part,
            new Inventory(),
            model.Flatten,
            !model.BeforeReferences,
            !model.BeforeSections,
            !model.BeforeComments,
            !model.BeforeRevisions,
            !model.BeforeNotes,
            !model.BeforeMath);
        var (_, blocks) = reader.Read(body);
        var index = blocks.ToDictionary(block => block.Oid, StringComparer.Ordinal);

        var section = body.Elements<SectionProperties>().LastOrDefault();

        // Parts outside `word/document.xml` this save may touch; it only grows when something
        // changes.
        var touched = new HashSet<string>(StringComparer.Ordinal);

        // Styles before the body, so a block finds a new style. In an old draft, only those the
        // package lacks (see StyleWriter.Apply).
        StyleWriter.Apply(part, model.Styles, inventory, touched, additionsOnly: model.Flatten);

        // Each note reference's id before the body, because the reference run carries it.
        var notes = model.BeforeNotes ? null : NotesWriter.Plan(model.Doc, part);

        var numbering = new NumberingFactory(part, touched, inventory);
        var headings = new HeadingStyles(part, touched);
        var replacement = BuildBody(
            model,
            part,
            index,
            inventory,
            numbering,
            headings,
            out var preserved,
            out var rewritten,
            out var breaks);

        body.RemoveAllChildren();
        foreach (var element in replacement) body.AppendChild(element);

        // Only the note that changed. Before comments, which also look for anchors in notes.
        if (notes is not null)
        {
            rewritten += NotesWriter.Apply(
                part,
                notes,
                reader.Notes,
                owner => new ParagraphWriter(
                    part,
                    inventory,
                    UsableWidthPx(model.Page),
                    headings,
                    model.Flatten,
                    !model.BeforeReferences,
                    !model.BeforeRevisions,
                    owner),
                numbering,
                inventory,
                touched,
                model.BeforeComments,
                model.BeforeRevisions,
                model.BeforeMath);
        }

        // Only when the model asks for numbering other than the package's.
        NotesWriter.ApplyNumbering(part, model.Notes, touched);

        // Before mending anchor ends, which needs to know the new comments.
        CommentsWriter.Apply(part, model, inventory, touched);

        // Ends the edit left unpaired, in the body and the notes; see MendCommentAnchors.
        var knownComments = (part.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? [])
            .Select(comment => comment.Id?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal);
        MendCommentAnchors(body, knownComments);
        foreach (OpenXmlPart? notesPart in new OpenXmlPart?[] { part.FootnotesPart, part.EndnotesPart })
        {
            if (notesPart?.RootElement is not OpenXmlPartRootElement notesRoot) continue;
            if (!MendCommentAnchors(notesRoot, knownComments)) continue;
            notesRoot.Save();
            touched.Add(notesPart.Uri.ToString().TrimStart('/'));
        }

        // A split move becomes a deletion and an insertion, each revision with its own `w:id`.
        MendMoves(body);
        Revisions.MakeIdsUnique(body, part);
        Revisions.ApplyTracking(part, model.TrackChanges, touched, inventory);

        body.AppendChild(section is null ? new SectionProperties() : section);

        // Only if the page changed: the model only knows A4 and Letter. See PageReader.Matches.
        var current = body.Elements<SectionProperties>().Last();
        if (!PageReader.Matches(current, model.Page)) ApplyPageSetup(current, model.Page);

        // Sections before the last, by the marks the body carried (see SectionWriter).
        var aliases = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!model.BeforeSections)
        {
            SectionWriter.ApplyStart(current, model.Page);
            SectionWriter.ApplyColumns(current, model.Page);
            aliases = SectionWriter.Apply(part, breaks, current, model, inventory, touched);
        }

        // The new document; in one from outside, the band wins. See PlainBandWriter.
        PlainBandWriter.Apply(part, current, model.Page, inventory, touched);

        PageNumbering.Apply(part, current, model.Page, touched, inventory);

        // Only band parts that changed enter the writable list.
        touched.UnionWith(BandWriter.Apply(part, [.. model.Sections ?? [], model.Page], inventory, aliases));

        // `docProps/custom.xml` is never touched.
        DocumentProperties.Apply(document, model.Properties, touched);

        part.Document!.Save();
        document.Dispose();

        // The destination content type and `.dotm` macros, on the final bytes.
        var restored = RestoreUntouchedParts(original, buffer.ToArray(), touched);
        return (PackageKind.Retarget(restored, model.Template, inventory),
            new SaveResult(inventory, preserved, rewritten));
    }

    /// <summary>
    /// Parts saving may change; a band part someone typed into enters separately, one by one, so a
    /// header nobody opened is not reserialized.
    /// </summary>
    private static readonly HashSet<string> Writable = new(StringComparer.Ordinal)
    {
        "word/document.xml",
        // An image or link inserted in an edited block.
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
    };

    /// <summary>
    /// The SDK **reserializes every part whose typed DOM was read**, even without changes: reading
    /// <c>NumberingDefinitionsPart.Numbering</c> is enough for <c>numbering.xml</c> to come out
    /// different. The invariant is enforced here.
    /// </summary>
    private static byte[] RestoreUntouchedParts(byte[] original, byte[] produced, HashSet<string> edited)
    {
        using var originalArchive = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
        var pristine = originalArchive.Entries.ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var stream = entry.Open();
                using var copy = new MemoryStream();
                stream.CopyTo(copy);
                return copy.ToArray();
            },
            StringComparer.Ordinal);

        using var source = new ZipArchive(new MemoryStream(produced), ZipArchiveMode.Read);
        using var result = new MemoryStream();

        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in source.Entries)
            {
                var keepOriginal = !Writable.Contains(entry.FullName) &&
                                   !edited.Contains(entry.FullName) &&
                                   pristine.ContainsKey(entry.FullName);

                using var target = output.CreateEntry(entry.FullName, CompressionLevel.Optimal).Open();

                if (keepOriginal)
                {
                    var bytes = pristine[entry.FullName];
                    target.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    using var stream = entry.Open();
                    stream.CopyTo(target);
                }
            }
        }

        return result.ToArray();
    }

    /// <summary>
    /// The editor deletes one end along with the text, and the preserved paragraph keeps its own;
    /// LibreOffice drops the thread or refuses the file. So:
    /// <list type="bullet">
    /// <item>a comment anchor the package does not have goes;</item>
    /// <item>a start without end or reference becomes a point comment;</item>
    /// <item>an end without a start goes, and the reference stays;</item>
    /// <item>a range without a reference gets one right after its end.</item>
    /// </list>
    /// Only zero-width elements: a preserved paragraph stays preserved.
    /// </summary>
    /// <returns>Whether anything changed; the notes part is only written then.</returns>
    private static bool MendCommentAnchors(OpenXmlElement body, HashSet<string> known)
    {
        var changed = false;
        var starts = body.Descendants<CommentRangeStart>().ToList();
        var ends = body.Descendants<CommentRangeEnd>().ToList();
        var references = body.Descendants<CommentReference>().ToList();

        static string IdOf(OpenXmlElement element) => element.GetAttribute("id", element.NamespaceUri).Value ?? string.Empty;
        void Remove(OpenXmlElement element)
        {
            changed = true;
            // Alone in its run, the reference takes the run along.
            if (element is CommentReference && element.Parent is Run run && BodyReader.ReferenceOnly(run) is not null)
            {
                run.Remove();
            }
            else
            {
                element.Remove();
            }
        }

        foreach (var element in starts.Concat<OpenXmlElement>(ends).Concat(references))
        {
            if (!known.Contains(IdOf(element))) Remove(element);
        }

        var startIds = starts.Where(start => start.Parent is not null).Select(IdOf).ToHashSet(StringComparer.Ordinal);
        var endIds = ends.Where(end => end.Parent is not null).Select(IdOf).ToHashSet(StringComparer.Ordinal);
        var referenceIds = references.Where(reference => reference.Parent is not null).Select(IdOf)
            .ToHashSet(StringComparer.Ordinal);

        static Run ReferenceRun(string id) =>
            new(new RunProperties(new RunStyle { Val = "CommentReference" }), new CommentReference { Id = id });

        foreach (var start in starts.Where(start => start.Parent is not null))
        {
            var id = IdOf(start);
            if (endIds.Contains(id)) continue;
            if (!referenceIds.Contains(id) && start.Parent is Paragraph or Hyperlink or SimpleField)
            {
                start.InsertAfterSelf(ReferenceRun(id));
                referenceIds.Add(id);
            }

            start.Remove();
            changed = true;
        }

        foreach (var end in ends.Where(end => end.Parent is not null))
        {
            var id = IdOf(end);
            if (!startIds.Contains(id))
            {
                end.Remove();
                changed = true;
            }
            else if (!referenceIds.Contains(id) && end.Parent is Paragraph or Hyperlink or SimpleField)
            {
                end.InsertAfterSelf(ReferenceRun(id));
                referenceIds.Add(id);
                changed = true;
            }
        }

        return changed;
    }

    /// <summary>
    /// A rewritten paragraph carries the moved range as <c>w:del</c>/<c>w:ins</c> and loses the
    /// ends; an incomplete pair becomes a deletion and an insertion on both sides, and loose ends
    /// go (the loss was already declared in NoteWhatWasInside).
    /// </summary>
    private static void MendMoves(Body body)
    {
        var fromStarts = body.Descendants<MoveFromRangeStart>().ToList();
        var toStarts = body.Descendants<MoveToRangeStart>().ToList();
        var fromEnds = body.Descendants<MoveFromRangeEnd>().ToList();
        var toEnds = body.Descendants<MoveToRangeEnd>().ToList();
        var moves = body.Descendants().Where(element => element is MoveFromRun or MoveToRun ||
            (element.Parent is ParagraphMarkRunProperties && element.LocalName is "moveFrom" or "moveTo")).ToList();
        if (fromStarts.Count + toStarts.Count + fromEnds.Count + toEnds.Count + moves.Count == 0) return;

        static string? Id(OpenXmlElement element) => Revisions.AttributeOf(element, "id");
        static string? Name(OpenXmlElement element) => Revisions.AttributeOf(element, "name");

        var fromEndIds = fromEnds.Select(Id).ToHashSet(StringComparer.Ordinal);
        var toEndIds = toEnds.Select(Id).ToHashSet(StringComparer.Ordinal);
        var fromNames = fromStarts.Where(start => fromEndIds.Contains(Id(start))).Select(Name)
            .ToHashSet(StringComparer.Ordinal);
        var toNames = toStarts.Where(start => toEndIds.Contains(Id(start))).Select(Name)
            .ToHashSet(StringComparer.Ordinal);
        var whole = fromNames.Intersect(toNames, StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);

        var keptIds = new HashSet<string?>(StringComparer.Ordinal);
        foreach (var start in fromStarts.Concat<OpenXmlElement>(toStarts))
        {
            if (whole.Contains(Name(start))) keptIds.Add((start is MoveFromRangeStart ? "f" : "t") + Id(start));
        }

        var inside = new HashSet<OpenXmlElement>(ReferenceEqualityComparer.Instance);
        var open = new List<string>();
        foreach (var element in body.Descendants())
        {
            switch (element)
            {
                case MoveFromRangeStart start when keptIds.Contains("f" + Id(start)): open.Add("f" + Id(start)); break;
                case MoveToRangeStart start when keptIds.Contains("t" + Id(start)): open.Add("t" + Id(start)); break;
                case MoveFromRangeEnd end: open.Remove("f" + Id(end)); break;
                case MoveToRangeEnd end: open.Remove("t" + Id(end)); break;
                default:
                    var side = element is MoveFromRun || element.LocalName == "moveFrom" ? "f" : "t";
                    if (moves.Contains(element) && open.Any(id => id.StartsWith(side, StringComparison.Ordinal)))
                    {
                        inside.Add(element);
                    }

                    break;
            }
        }

        foreach (var marker in fromStarts.Concat<OpenXmlElement>(toStarts))
        {
            if (!whole.Contains(Name(marker))) marker.Remove();
        }

        foreach (var end in fromEnds)
        {
            if (!keptIds.Contains("f" + Id(end))) end.Remove();
        }

        foreach (var end in toEnds)
        {
            if (!keptIds.Contains("t" + Id(end))) end.Remove();
        }

        foreach (var move in moves.Where(move => !inside.Contains(move)))
        {
            OpenXmlElement plain = move switch
            {
                MoveFromRun => new DeletedRun(),
                MoveToRun => new InsertedRun(),
                _ when move.LocalName == "moveFrom" => new Deleted(),
                _ => new Inserted(),
            };

            Revisions.Stamp(
                plain,
                Revisions.AttributeOf(move, "id"),
                Revisions.AttributeOf(move, "author"),
                Revisions.AttributeOf(move, "date"));
            foreach (var child in move.ChildElements.ToList())
            {
                child.Remove();
                plain.AppendChild(child);
            }

            move.InsertAfterSelf(plain);
            move.Remove();
        }
    }

    private static List<OpenXmlElement> BuildBody(
        DocumentModelDto model,
        MainDocumentPart part,
        Dictionary<string, Block> index,
        Inventory inventory,
        NumberingFactory numbering,
        HeadingStyles headings,
        out int preserved,
        out int rewritten,
        out List<SectionWriter.Break> breaks)
    {
        var writer = new ParagraphWriter(
            part,
            inventory,
            UsableWidthPx(model.Page),
            headings,
            model.Flatten,
            !model.BeforeReferences,
            !model.BeforeRevisions);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var elements = new List<OpenXmlElement>();
        var generated = new HashSet<OpenXmlElement>(ReferenceEqualityComparer.Instance);

        preserved = 0;
        rewritten = 0;
        breaks = [];
        var breakIds = new HashSet<string>(StringComparer.Ordinal);
        var knownSections = (model.Sections ?? []).Select(section => section.Id).OfType<string>()
            .ToHashSet(StringComparer.Ordinal);

        foreach (var slot in Flatten(model.Doc, numbering))
        {
            var oid = OidOf(slot.Identity);

            // A repeated `oid` is a pasted block: the original XML belongs to **one** of them. From
            // the second occurrence on the block is new.
            var first = oid is not null && used.Add(oid);

            // What the body kept before this block goes back before it (Block.Leading), only on the
            // first occurrence.
            var owner = first && index.TryGetValue(oid!, out var known) ? known : null;
            if (owner is not null)
            {
                foreach (var loose in owner.Leading) elements.Add(loose.CloneNode(true));
            }

            var before = elements.Count;
            var kept = BuildSlot(
                slot, owner, writer, inventory, elements, model.BeforeComments, model.BeforeRevisions, model.BeforeNotes,
                model.BeforeMath);
            if (kept)
            {
                preserved++;
            }
            else
            {
                rewritten++;
                for (var i = before; i < elements.Count; i++) generated.Add(elements[i]);
            }

            if (!model.BeforeSections &&
                SectionWriter.Mark(
                    slot.Content, elements.Skip(before).OfType<Paragraph>().ToList(), kept, breakIds, knownSections, inventory)
                    is { } mark)
            {
                breaks.Add(mark);
            }

            if (owner is not null)
            {
                foreach (var loose in owner.Trailing) elements.Add(loose.CloneNode(true));
            }
        }

        // A loose bookmark goes with the deleted block, as in Word; nothing else is deleted
        // silently.
        foreach (var block in index.Values.Where(block => !used.Contains(block.Oid)))
        {
            if (block.Leading.Concat(block.Trailing).Any(loose => loose is not (BookmarkStart or BookmarkEnd)))
            {
                inventory.NoteLoss("conteúdo solto entre blocos que você apagou");
            }
        }

        UniqueBookmarks(elements, generated);
        MatchLooseBookmarks(elements, index.Values.Where(block => !used.Contains(block.Oid)));

        if (elements.Count == 0) elements.Add(new Paragraph());
        return elements;
    }

    /// <summary>
    /// Word writes between paragraphs the end of a bookmark that ends after a table. Once the block
    /// holding it is deleted, the remaining end goes back next to the other's block, and the
    /// bookmark shrinks to what is left of it, as in Word.
    /// </summary>
    private static void MatchLooseBookmarks(List<OpenXmlElement> elements, IEnumerable<Block> deleted)
    {
        foreach (var loose in deleted.SelectMany(block => block.Leading.Concat(block.Trailing)))
        {
            if (loose is BookmarkEnd end && end.Id?.Value is { } endId)
            {
                var owner = elements.FindIndex(element => Has<BookmarkStart>(element, endId));
                if (owner >= 0 && !elements.Any(element => Has<BookmarkEnd>(element, endId)))
                {
                    elements.Insert(owner + 1, end.CloneNode(true));
                }
            }
            else if (loose is BookmarkStart start && start.Id?.Value is { } startId)
            {
                var owner = elements.FindIndex(element => Has<BookmarkEnd>(element, startId));
                if (owner >= 0 && !elements.Any(element => Has<BookmarkStart>(element, startId)))
                {
                    elements.Insert(owner, start.CloneNode(true));
                }
            }
        }

        static bool Has<T>(OpenXmlElement element, string id)
            where T : OpenXmlElement =>
            element.Descendants<T>().Cast<OpenXmlElement>().Prepend(element).Where(mark => mark is T).Any(mark =>
                ((mark as BookmarkStart)?.Id?.Value ?? (mark as BookmarkEnd)?.Id?.Value) == id);
    }

    /// <summary>
    /// Word refuses two <c>w:bookmarkStart</c>s with the same id. The block generated now yields;
    /// between two generated ones the first stays, as Word resolves it.
    /// </summary>
    private static void UniqueBookmarks(List<OpenXmlElement> elements, HashSet<OpenXmlElement> generated)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ends = new HashSet<string>(StringComparer.Ordinal);

        // The highest id in the whole body, including those the model does not know.
        var highest = elements
            .SelectMany(element => Starts(element).Select(start => start.Id?.Value)
                .Concat(Ends(element).Select(end => end.Id?.Value)))
            .Select(id => int.TryParse(id, out var value) ? value : -1)
            .DefaultIfEmpty(-1)
            .Max();

        foreach (var kept in elements.Where(element => !generated.Contains(element)))
        {
            foreach (var start in Starts(kept))
            {
                ids.Add(start.Id?.Value ?? string.Empty);
                names.Add(start.Name?.Value ?? string.Empty);
            }

            foreach (var end in Ends(kept)) ends.Add(end.Id?.Value ?? string.Empty);
        }

        var dropped = new HashSet<string>(StringComparer.Ordinal);
        var renamed = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var element in elements.Where(generated.Contains))
        {
            // In document order, so the end finds the renumbered start.
            foreach (var mark in element.Descendants().Prepend(element).ToList())
            {
                if (mark is BookmarkStart start)
                {
                    var id = start.Id?.Value ?? string.Empty;
                    var name = start.Name?.Value ?? string.Empty;

                    // A repeated name is a pasted copy: it goes, with its end.
                    if (!names.Add(name))
                    {
                        dropped.Add(id);
                        start.Remove();
                        continue;
                    }

                    // Only a repeated id comes from another document: it gets a new id.
                    if (!ids.Add(id))
                    {
                        var fresh = (++highest).ToString(System.Globalization.CultureInfo.InvariantCulture);
                        renamed[id] = fresh;
                        start.Id = fresh;
                        ids.Add(fresh);
                    }
                }
                else if (mark is BookmarkEnd end)
                {
                    var id = end.Id?.Value ?? string.Empty;
                    if (renamed.Remove(id, out var fresh))
                    {
                        end.Id = fresh;
                        ends.Add(fresh);
                        continue;
                    }

                    if (dropped.Remove(id) || !ends.Add(id)) end.Remove();
                }
            }
        }

        static IEnumerable<BookmarkStart> Starts(OpenXmlElement element) =>
            element is BookmarkStart start ? [start] : element.Descendants<BookmarkStart>();

        static IEnumerable<BookmarkEnd> Ends(OpenXmlElement element) =>
            element is BookmarkEnd end ? [end] : element.Descendants<BookmarkEnd>();
    }

    /// <summary>The original XML when the content did not change, the rewritten one when it did.
    /// Returns whether it was preserved.</summary>
    /// <param name="owner">The file block with the same `oid`, on its first occurrence.</param>
    internal static bool BuildSlot(
        Slot slot,
        Block? owner,
        ParagraphWriter writer,
        Inventory inventory,
        List<OpenXmlElement> elements,
        bool beforeComments,
        bool beforeRevisions,
        bool beforeNotes,
        bool beforeMath)
    {
        if (owner is not null && SameContent(slot, owner))
        {
            // Tab and "Restart numbering" change the surrounding list, not the item: only `w:numPr`
            // is replaced, and the rest goes back as it came.
            if (slot.List is { } list && owner.Source is Paragraph paragraph && !Points(paragraph, list))
            {
                elements.Add(Renumbered(paragraph, list));
                return false;
            }

            elements.Add(owner.Source.CloneNode(true));
            foreach (var next in owner.Continuation) elements.Add(next.CloneNode(true));
            return true;
        }

        // The original still gives the rewritten paragraph the anchored objects this writer does
        // not generate.
        var source = owner?.Source;

        // Here it is known whether the paragraph is a list item; in a cell, it is not
        // (ParagraphWriter.ListPlacement).
        var placement = new ParagraphWriter.ListPlacement(slot.List);

        foreach (var element in writer.Write(slot.Content, placement, source)) elements.Add(element);

        if (owner is not null) NoteWhatWasInside(owner, inventory, beforeComments, beforeRevisions, beforeNotes, beforeMath);
        return false;
    }

    /// <summary>See OwnContent.</summary>
    private static bool SameContent(Slot slot, Block owner) =>
        string.Equals(
            OwnContent(owner.Extracted).Fingerprint(),
            OwnContent(slot.Identity).Fingerprint(),
            StringComparison.Ordinal);

    /// <summary>The same content, and a list item in the same numbering; see BuildSlot.</summary>
    internal static bool Preservable(Slot slot, Block owner) =>
        SameContent(slot, owner) &&
        !(slot.List is { } list && owner.Source is Paragraph paragraph && !Points(paragraph, list));

    /// <summary>
    /// The item without its sublists: in the file they are the following paragraphs, and compared
    /// with them the item above would be rewritten on every Tab in one below.
    /// </summary>
    private static Node OwnContent(Node node)
    {
        if (node.Type != "listItem" || node.Content is null) return node;
        return new Node
        {
            Type = node.Type,
            Attrs = node.Attrs,
            Content = [.. node.Content.Where(child => child.Type is not ("bulletList" or "orderedList"))],
        };
    }

    private static bool Points(Paragraph paragraph, ParagraphWriter.ListContext list)
    {
        var numbering = paragraph.ParagraphProperties?.NumberingProperties;
        return numbering?.NumberingId?.Val?.Value == list.NumberingId &&
               (numbering.NumberingLevelReference?.Val?.Value ?? 0) == list.Level;
    }

    private static Paragraph Renumbered(Paragraph original, ParagraphWriter.ListContext list)
    {
        var paragraph = (Paragraph)original.CloneNode(true);
        var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        properties.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = list.Level },
            new NumberingId { Val = list.NumberingId });
        return paragraph;
    }

    /// <summary>
    /// In the editor a list has items inside; in OOXML they are sibling paragraphs, and the
    /// <c>oid</c> lives on the <c>listItem</c>.
    /// </summary>
    /// <param name="Identity">
    /// The node extracted on open: for an item, the <c>listItem</c>, not the paragraph, or nothing
    /// would be preserved.
    /// </param>
    internal sealed record Slot(Node Identity, Node Content, ParagraphWriter.ListContext? List);

    internal static IEnumerable<Slot> Flatten(
        Node doc,
        NumberingFactory numbering,
        ParagraphWriter.ListContext? inherited = null)
    {
        foreach (var node in doc.Content ?? [])
        {
            switch (node.Type)
            {
                case "bulletList":
                case "orderedList":
                {
                    // The file's level; otherwise, one below the outer list.
                    var level = Math.Clamp(
                        Attr.Int(node, "level") ?? (inherited?.Level ?? -1) + 1, 0, ListLevels.Count - 1);
                    var definition = Attr.Node(node, "numbering") as System.Text.Json.Nodes.JsonObject;

                    // The outer numbering only serves an inner list of the same kind, and only when
                    // the inner one did not bring its own definition, the restart's.
                    var fromParent = inherited is { } outer
                                     && definition is null
                                     && string.Equals(outer.Kind, node.Type, StringComparison.Ordinal)
                        ? (int?)outer.NumberingId
                        : null;

                    // The file's `numId`, the outer list's, or a new definition (a list born here
                    // or pasted from another document). Never zero, which in the format means "no
                    // numbering".
                    var numberingId = numbering.NumberingIdFor(
                        node.Type, NumberingOf(node) ?? fromParent, definition);

                    foreach (var item in node.Content ?? [])
                    {
                        if (item.Type != "listItem") continue;

                        var context = new ParagraphWriter.ListContext(node.Type, numberingId, level);
                        var paragraphs = (item.Content ?? [])
                            .Where(child => child.Type is "paragraph" or "heading").ToList();
                        var nested = (item.Content ?? [])
                            .Where(child => child.Type is "bulletList" or "orderedList");

                        if (paragraphs.Count > 0) yield return new Slot(item, paragraphs[0], context);

                        foreach (var extra in paragraphs.Skip(1))
                        {
                            yield return new Slot(extra, extra, context);
                        }

                        foreach (var child in nested)
                        {
                            var wrapper = Node.Of("doc");
                            wrapper.Content = [child];
                            foreach (var deeper in Flatten(wrapper, numbering, context)) yield return deeper;
                        }
                    }

                    break;
                }

                case "blockquote":
                {
                    // No quote in OOXML: it becomes an indent, as in Word.
                    foreach (var child in node.Content ?? [])
                    {
                        yield return new Slot(child, child, inherited);
                    }

                    break;
                }

                default:
                    yield return new Slot(node, node, inherited);
                    break;
            }
        }
    }

    private static int? NumberingOf(Node list) =>
        Attr.Int(list, "numId") is { } numId && numId > 0 ? numId : null;

    /// <summary>
    /// The ceiling for an image without a measure: a 1920 px screenshot would pass both margins.
    /// </summary>
    private static int UsableWidthPx(PageSetupDto page)
    {
        var (shortSide, longSide) = PageReader.MillimetersOfPaper(page.Size);

        var across = string.Equals(page.Orientation, "landscape", StringComparison.Ordinal)
            ? longSide
            : shortSide;

        var millimeters = across - page.Margins.Left - page.Margins.Right;
        return millimeters > 10 ? (int)Math.Round(millimeters / Unit.MillimetersPerInch * Unit.PixelsPerInch) : ImageWriter.DefaultWidthPx;
    }

    internal static string? OidOf(Node node)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue("oid", out var value) || value is null) return null;
        return value.GetValueKind() == System.Text.Json.JsonValueKind.String ? value.GetValue<string>() : null;
    }

    /// <summary>
    /// The **real loss**, detected by comparison: the warning says "you edited a paragraph that had
    /// an anchored comment". Each <c>before…</c> is the draft older than the feature, whose nodes
    /// do not carry it and whose rewritten paragraph loses it.
    /// </summary>
    private static void NoteWhatWasInside(
        Block block, Inventory inventory, bool beforeComments, bool beforeRevisions, bool beforeNotes, bool beforeMath)
    {
        var original = block.Source;

        if (beforeComments
                ? original.Descendants<CommentRangeStart>().Any() || original.Descendants<CommentReference>().Any()
                : original.Descendants<CommentReference>().Any(reference =>
                    reference.Parent is Run run && BodyReader.ReferenceOnly(run) is null))
        {
            inventory.NoteLoss("comentário ancorado num parágrafo que você editou");
        }

        // In an old draft moves and formatting revisions are lost too.
        if (beforeRevisions &&
            (original.Descendants<InsertedRun>().Any() || original.Descendants<DeletedRun>().Any() ||
             original.Descendants<MoveFromRun>().Any() || original.Descendants<MoveToRun>().Any() ||
             original.Descendants<MoveFromRangeStart>().Any() || original.Descendants<MoveToRangeStart>().Any() ||
             original.Descendants<RunPropertiesChange>().Any(change => change.Parent?.Parent is Run)))
        {
            inventory.NoteLoss("marcas de revisão num parágrafo que você editou");
        }

        // An image inside a revision does not carry the mark: it goes back as accepted content.
        if (!beforeRevisions &&
            original.Descendants().Any(element =>
                element is InsertedRun or DeletedRun or MoveFromRun or MoveToRun &&
                (element.Descendants<Drawing>().Any() || element.Descendants<Picture>().Any() ||
                 element.Descendants<EmbeddedObject>().Any())))
        {
            inventory.NoteLoss("imagem dentro de uma revisão num parágrafo que você editou (ficou como aceita)");
        }

        // Formatting from before the revision lives in `w:rPr`, which saving rebuilds from the
        // marks.
        if (!beforeRevisions && original.Descendants<RunPropertiesChange>().Any(change => change.Parent?.Parent is Run))
        {
            inventory.NoteLoss("revisão de formatação num parágrafo que você editou");
        }

        if (!beforeRevisions &&
            (original.Descendants<MoveFromRangeStart>().Any() || original.Descendants<MoveToRangeStart>().Any() ||
             original.Descendants<MoveFromRangeEnd>().Any() || original.Descendants<MoveToRangeEnd>().Any()))
        {
            inventory.NoteLoss("movimentação de texto num parágrafo que você editou (virou exclusão e inserção)");
        }

        if (beforeNotes && original.Descendants<FootnoteReference>().Any())
        {
            inventory.NoteLoss("nota de rodapé num parágrafo que você editou");
        }

        if (beforeNotes && original.Descendants<EndnoteReference>().Any())
        {
            inventory.NoteLoss("nota de fim num parágrafo que você editou");
        }

        if (beforeMath && original.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any())
        {
            inventory.NoteLoss("equação num parágrafo que você editou");
        }

        // A field that became a node goes back as a field; one shown only by its result is lost.
        if (block.UnrepresentedField)
        {
            inventory.NoteLoss("campo calculado num parágrafo que você editou");
        }

        // Anchored objects are copied from the original; what is left is VML (`w:pict`) and a
        // drawing in a run with text.
        if (original.Descendants<Picture>().Any() ||
            original.Elements<Run>().Any(run =>
                run.Descendants<WordDrawing.Anchor>().Any() && !ParagraphWriter.IsAnchoredOnly(run)))
        {
            inventory.NoteLoss("forma ou caixa de texto num parágrafo que você editou");
        }
    }

    internal static void ApplyPageSetup(SectionProperties section, PageSetupDto page)
    {
        var landscape = string.Equals(page.Orientation, "landscape", StringComparison.Ordinal);
        var (shortSide, longSide) = PageReader.TwipsOfPaper(page.Size);

        var size = section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
        if (size is null)
        {
            size = new DocumentFormat.OpenXml.Wordprocessing.PageSize();
            section.PrependChild(size);
        }
        else if (PreservedPaper(size, page.Size) is { } measured)
        {
            // The paper stays what the model says: an A5 does not become A4 over a margin change.
            (shortSide, longSide) = measured;
        }

        size.Width = landscape ? longSide : shortSide;
        size.Height = landscape ? shortSide : longSide;
        size.Orient = landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;

        var margin = section.GetFirstChild<PageMargin>();
        if (margin is null)
        {
            margin = new PageMargin();
            section.InsertAfter(margin, size);
        }

        margin.Top = Attr.MmToTwips(page.Margins.Top);
        margin.Bottom = Attr.MmToTwips(page.Margins.Bottom);
        margin.Left = (uint)Math.Max(0, Attr.MmToTwips(page.Margins.Left));
        margin.Right = (uint)Math.Max(0, Attr.MmToTwips(page.Margins.Right));
    }

    /// <summary>When they still match the paper the model names.</summary>
    private static (uint Short, uint Long)? PreservedPaper(
        DocumentFormat.OpenXml.Wordprocessing.PageSize size,
        string wanted)
    {
        if (size.Width?.Value is not { } width || size.Height?.Value is not { } height) return null;
        if (width == 0 || height == 0) return null;

        var landscape = size.Orient is not null && size.Orient.Value == PageOrientationValues.Landscape;
        return string.Equals(PageReader.NameOfPaper(width, height, landscape), wanted, StringComparison.Ordinal)
            ? (Math.Min(width, height), Math.Max(width, height))
            : null;
    }

    private static WordprocessingDocument OpenEditable(Stream stream)
    {
        try
        {
            return WordprocessingDocument.Open(stream, isEditable: true);
        }
        catch (Exception problem) when (problem is not DocxException)
        {
            throw new DocxException(
                "Não foi possível gravar sobre o arquivo original. Ele pode ter sido alterado ou danificado.",
                problem);
        }
    }
}

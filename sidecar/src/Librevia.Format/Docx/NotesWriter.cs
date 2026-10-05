using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A note whose body did not change is not touched, and a part without a changed note goes back
/// byte for byte; in an edited note, only the edited paragraph is rewritten. The number is the
/// reference's order; <c>nid</c> is the <c>w:id</c>. A new or repeated note gets an id above the
/// highest; one that lost its reference goes, as in Word.
/// </summary>
internal static class NotesWriter
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public const string Footnote = "footnote";
    public const string Endnote = "endnote";

    /// <c>fn:3</c>, <c>en:1</c>; its blocks are <c>fn:3/p1</c>…
    public static string Address(bool endnote, string id) => (endnote ? "en:" : "fn:") + id;

    /// <param name="Fresh">The note will be created with <paramref name="Id"/>.</param>
    /// <param name="Original">The id it had in the file, when renumbering changed it.</param>
    public sealed record Wanted(Node Reference, bool Endnote, string Id, bool Fresh, string? Original = null)
    {
        public string Source => Original ?? Id;
    }

    /// <summary>The normal note, not the separator.</summary>
    public static OpenXmlElement? NoteOf(MainDocumentPart part, bool endnote, string id) =>
        NotesIn(part, endnote).FirstOrDefault(note => IsNormal(note) && IdOf(note) == id);

    private static IEnumerable<FootnoteEndnoteType> NotesIn(MainDocumentPart part, bool endnote) =>
        endnote
            ? part.EndnotesPart?.Endnotes?.Elements<DocumentFormat.OpenXml.Wordprocessing.Endnote>() ?? []
            : (IEnumerable<FootnoteEndnoteType>?)part.FootnotesPart?.Footnotes?.Elements<DocumentFormat.OpenXml.Wordprocessing.Footnote>() ?? [];

    private static bool IsNormal(FootnoteEndnoteType note) =>
        note.Type?.Value is null || note.Type.Value == FootnoteEndnoteValues.Normal;

    private static string? IdOf(FootnoteEndnoteType note) =>
        note.Id?.Value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Before the body, because the reference run carries the id; a paragraph with a new id is
    /// rewritten.
    /// </summary>
    public static List<Wanted> Plan(Node doc, MainDocumentPart part)
    {
        var wanted = new List<Wanted>();
        var used = new HashSet<string>(StringComparer.Ordinal);
        var highest = new Dictionary<bool, long>
        {
            [false] = HighestId(part, endnote: false),
            [true] = HighestId(part, endnote: true),
        };

        foreach (var reference in ReferencesIn(doc))
        {
            var endnote = Attr.String(reference, "kind") == Endnote;
            var nid = Attr.String(reference, "nid");
            if (nid is not null && NoteOf(part, endnote, nid) is not null && used.Add(Address(endnote, nid)))
            {
                wanted.Add(new Wanted(reference, endnote, nid, Fresh: false));
                continue;
            }

            var fresh = (highest[endnote] = highest[endnote] + 1).ToString(CultureInfo.InvariantCulture);
            used.Add(Address(endnote, fresh));
            reference.With("nid", fresh);
            wanted.Add(new Wanted(reference, endnote, fresh, Fresh: true));
        }

        return InTextOrder(wanted, part);
    }

    /// <summary>
    /// LibreOffice assigns notes to references in id order: a moved note would show with another's
    /// body. A file already out of order and a save that moved nothing stay as they are.
    /// </summary>
    private static List<Wanted> InTextOrder(List<Wanted> wanted, MainDocumentPart part)
    {
        var result = new List<Wanted>(wanted);
        foreach (var endnote in new[] { false, true })
        {
            if (!Ascending(OriginalOrder(part, endnote))) continue;
            var mine = result.Select((entry, index) => (entry, index)).Where(pair => pair.entry.Endnote == endnote).ToList();
            // A new note gets the highest id and breaks the order like a moved one.
            if (Ascending(mine.Select(pair => pair.entry.Id))) continue;

            var next = NotesIn(part, endnote).Where(note => !IsNormal(note))
                .Select(note => note.Id?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
            foreach (var (entry, index) in mine)
            {
                var id = (next++).ToString(CultureInfo.InvariantCulture);
                entry.Reference.With("nid", id);
                result[index] = entry with { Id = id, Original = entry.Fresh ? null : entry.Id };
            }
        }

        return result;
    }

    private static bool Ascending(IEnumerable<string> ids)
    {
        long previous = long.MinValue;
        foreach (var id in ids)
        {
            if (!long.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value <= previous)
                return false;
            previous = value;
        }

        return true;
    }

    private static IEnumerable<string> OriginalOrder(MainDocumentPart part, bool endnote) =>
        (part.Document?.Body?.Descendants<Run>() ?? [])
            .Select(BodyReader.NoteReferenceOf)
            .OfType<FootnoteEndnoteReferenceType>()
            .Where(reference => reference is EndnoteReference == endnote)
            .Select(reference => reference.Id?.Value.ToString(CultureInfo.InvariantCulture))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal);

    private static long HighestId(MainDocumentPart part, bool endnote) =>
        NotesIn(part, endnote).Select(note => note.Id?.Value ?? 0L).DefaultIfEmpty(0L).Max() is var max && max > 0
            ? max
            : 0L;

    /// <summary>Without descending into any body.</summary>
    private static IEnumerable<Node> ReferencesIn(Node node)
    {
        foreach (var child in node.Content ?? [])
        {
            if (child.Type == "noteRef")
            {
                yield return child;
                continue;
            }

            foreach (var deeper in ReferencesIn(child)) yield return deeper;
        }
    }

    /// <summary>Returns how many note blocks were rewritten.</summary>
    /// <param name="read">The notes from the reference reading; see BodyReader.Notes.</param>
    /// <param name="writerFor">The writer owning each part's relationships.</param>
    public static int Apply(
        MainDocumentPart part,
        List<Wanted> wanted,
        IReadOnlyDictionary<string, BodyReader.NoteRead> read,
        Func<OpenXmlPart, ParagraphWriter> writerFor,
        NumberingFactory numbering,
        Inventory inventory,
        HashSet<string> touched,
        bool beforeComments,
        bool beforeRevisions,
        bool beforeMath)
    {
        var rewritten = 0;
        foreach (var endnote in new[] { false, true })
        {
            var mine = wanted.Where(entry => entry.Endnote == endnote).ToList();
            var prefix = endnote ? "en:" : "fn:";
            var gone = read
                .Where(entry => entry.Key.StartsWith(prefix, StringComparison.Ordinal) &&
                                !mine.Any(kept => !kept.Fresh && Address(endnote, kept.Source) == entry.Key))
                .Select(entry => entry.Value.Source)
                .ToList();

            // LibreOffice matches by order in the part: a moved note changes place there too,
            // compared with the original body order.
            var inOrder = read.Keys
                .Where(key => key.StartsWith(prefix, StringComparison.Ordinal))
                .Select(key => key[prefix.Length..])
                .Where(id => mine.Any(entry => !entry.Fresh && entry.Source == id))
                .SequenceEqual(mine.Where(entry => !entry.Fresh).Select(entry => entry.Source));

            if (inOrder && gone.Count == 0 && mine.All(entry => !entry.Fresh && entry.Original is null) &&
                mine.All(entry => !read.ContainsKey(Address(endnote, entry.Id)) ||
                                  Unchanged(read[Address(endnote, entry.Id)], entry.Reference, numbering)))
            {
                continue;
            }

            var owner = PartFor(part, endnote, touched);
            var root = owner.RootElement!;
            var writer = writerFor(owner);
            var mark = ReferenceMarkRun(root, part, endnote);

            foreach (var entry in mine)
            {
                if (!entry.Fresh)
                {
                    if (entry.Original is not null &&
                        read.TryGetValue(Address(endnote, entry.Source), out var renumbered) &&
                        renumbered.Source is FootnoteEndnoteType moved)
                    {
                        moved.Id = long.Parse(entry.Id, CultureInfo.InvariantCulture);
                    }

                    if (read.TryGetValue(Address(endnote, entry.Source), out var original) &&
                        !Unchanged(original, entry.Reference, numbering))
                    {
                        rewritten += Rebuild(original, entry.Reference, writer, numbering, inventory,
                            beforeComments, beforeRevisions, beforeMath);
                        EnsureReferenceMark(original.Source, mark, endnote);
                    }

                    continue;
                }

                var note = endnote
                    ? (FootnoteEndnoteType)new DocumentFormat.OpenXml.Wordprocessing.Endnote()
                    : new DocumentFormat.OpenXml.Wordprocessing.Footnote();
                note.Id = long.Parse(entry.Id, CultureInfo.InvariantCulture);
                var empty = new BodyReader.NoteRead(note, []);
                rewritten += Rebuild(empty, entry.Reference, writer, numbering, inventory, beforeComments, beforeRevisions, beforeMath);
                StyleNewNote(note, part, endnote);
                EnsureReferenceMark(note, mark, endnote);
                root.AppendChild(note);
            }

            foreach (var note in gone) note.Remove();

            var normals = root.Elements<FootnoteEndnoteType>().Where(IsNormal)
                .Where(note => IdOf(note) is not null)
                .GroupBy(note => IdOf(note)!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var entry in mine)
            {
                if (!normals.TryGetValue(entry.Id, out var note)) continue;
                note.Remove();
                root.AppendChild(note);
            }

            root.Save();
            var path = owner.Uri.ToString().TrimStart('/');
            touched.Add(path);
            touched.Add(RelationshipsPathOf(path));
        }

        return rewritten;
    }

    private static string RelationshipsPathOf(string path)
    {
        var slash = path.LastIndexOf('/');
        return slash < 0 ? $"_rels/{path}.rels" : $"{path[..slash]}/_rels/{path[(slash + 1)..]}.rels";
    }

    /// <summary>The same blocks, in the same order, with the same fingerprint.</summary>
    private static bool Unchanged(BodyReader.NoteRead original, Node reference, NumberingFactory numbering)
    {
        var doc = Node.Of("doc");
        doc.Content = reference.Content ?? [];
        var slots = DocxWriter.Flatten(doc, numbering).ToList();
        if (slots.Count != original.Blocks.Count) return false;

        for (var at = 0; at < slots.Count; at++)
        {
            var block = original.Blocks[at];
            if (DocxWriter.OidOf(slots[at].Identity) != block.Oid || !DocxWriter.Preservable(slots[at], block)) return false;
        }

        return true;
    }

    /// <summary>Returns how many blocks were rewritten.</summary>
    private static int Rebuild(
        BodyReader.NoteRead original,
        Node reference,
        ParagraphWriter writer,
        NumberingFactory numbering,
        Inventory inventory,
        bool beforeComments,
        bool beforeRevisions,
        bool beforeMath)
    {
        var index = original.Blocks.ToDictionary(block => block.Oid, StringComparer.Ordinal);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var elements = new List<OpenXmlElement>();
        var rewritten = 0;

        var doc = Node.Of("doc");
        doc.Content = reference.Content ?? [];
        foreach (var slot in DocxWriter.Flatten(doc, numbering))
        {
            var oid = DocxWriter.OidOf(slot.Identity);
            var first = oid is not null && used.Add(oid);
            var owner = first && index.TryGetValue(oid!, out var known) ? known : null;
            if (owner is not null)
            {
                foreach (var loose in owner.Leading) elements.Add(loose.CloneNode(true));
            }

            if (!DocxWriter.BuildSlot(slot, owner, writer, inventory, elements, beforeComments, beforeRevisions, beforeNotes: false, beforeMath))
            {
                rewritten++;
            }

            if (owner is not null)
            {
                foreach (var loose in owner.Trailing) elements.Add(loose.CloneNode(true));
            }
        }

        // What a deleted note carried between blocks goes with it.
        if (elements.Count == 0) elements.Add(new Paragraph());

        var source = original.Source;
        source.RemoveAllChildren();
        foreach (var element in elements) source.AppendChild(element);
        return rewritten;
    }

    /// <summary>
    /// The first paragraph must carry <c>w:footnoteRef</c>, which the reader did not give the
    /// model.
    /// </summary>
    private static void EnsureReferenceMark(OpenXmlElement note, Run mark, bool endnote)
    {
        var paragraph = note.Elements<Paragraph>().FirstOrDefault();
        if (paragraph is null) return;

        var present = endnote
            ? paragraph.Descendants<EndnoteReferenceMark>().Any()
            : paragraph.Descendants<FootnoteReferenceMark>().Any();
        if (present) return;

        var run = (Run)mark.CloneNode(true);
        if (paragraph.ParagraphProperties is { } properties) properties.InsertAfterSelf(run);
        else paragraph.PrependChild(run);
    }

    /// <summary>The part's, or one in Word's style.</summary>
    private static Run ReferenceMarkRun(OpenXmlElement root, MainDocumentPart part, bool endnote)
    {
        var existing = root.Descendants<Run>().FirstOrDefault(run =>
            endnote ? run.Elements<EndnoteReferenceMark>().Any() : run.Elements<FootnoteReferenceMark>().Any());
        if (existing is not null)
        {
            var copy = new Run();
            if (existing.RunProperties is { } properties) copy.RunProperties = (RunProperties)properties.CloneNode(true);
            copy.AppendChild(endnote ? new EndnoteReferenceMark() : (OpenXmlElement)new FootnoteReferenceMark());
            return copy;
        }

        return new Run(
            ReferenceProperties(part, endnote),
            endnote ? new EndnoteReferenceMark() : new FootnoteReferenceMark());
    }

    /// <summary>
    /// Word's character style when the document defines it; otherwise, direct superscript.
    /// </summary>
    public static RunProperties ReferenceProperties(MainDocumentPart part, bool endnote)
    {
        var style = endnote ? "EndnoteReference" : "FootnoteReference";
        return StyleDefined(part, style)
            ? new RunProperties(new RunStyle { Val = style })
            : new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
    }

    private static void StyleNewNote(OpenXmlElement note, MainDocumentPart part, bool endnote)
    {
        var style = endnote ? "EndnoteText" : "FootnoteText";
        if (!StyleDefined(part, style)) return;

        foreach (var paragraph in note.Elements<Paragraph>())
        {
            var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.ParagraphStyleId ??= new ParagraphStyleId { Val = style };
        }
    }

    private static bool StyleDefined(MainDocumentPart part, string styleId) =>
        part.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Any(style => string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal)) == true;

    /// <summary>
    /// Created, with the separators Word requires, when the document has no note of this kind.
    /// </summary>
    private static OpenXmlPart PartFor(MainDocumentPart part, bool endnote, HashSet<string> touched)
    {
        if (endnote && part.EndnotesPart is { Endnotes: not null } endnotes) return endnotes;
        if (!endnote && part.FootnotesPart is { Footnotes: not null } footnotes) return footnotes;

        OpenXmlPart created;
        if (endnote)
        {
            var owner = part.EndnotesPart ?? part.AddNewPart<EndnotesPart>();
            owner.Endnotes = new Endnotes(
                Separator<DocumentFormat.OpenXml.Wordprocessing.Endnote>(-1, FootnoteEndnoteValues.Separator, new SeparatorMark()),
                Separator<DocumentFormat.OpenXml.Wordprocessing.Endnote>(0, FootnoteEndnoteValues.ContinuationSeparator, new ContinuationSeparatorMark()));
            created = owner;
        }
        else
        {
            var owner = part.FootnotesPart ?? part.AddNewPart<FootnotesPart>();
            owner.Footnotes = new Footnotes(
                Separator<DocumentFormat.OpenXml.Wordprocessing.Footnote>(-1, FootnoteEndnoteValues.Separator, new SeparatorMark()),
                Separator<DocumentFormat.OpenXml.Wordprocessing.Footnote>(0, FootnoteEndnoteValues.ContinuationSeparator, new ContinuationSeparatorMark()));
            created = owner;
        }

        DeclareSeparators(part, endnote, touched);
        return created;
    }

    private static T Separator<T>(long id, FootnoteEndnoteValues type, OpenXmlElement mark)
        where T : FootnoteEndnoteType, new()
    {
        var note = new T { Type = type, Id = id };
        note.AppendChild(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new Run(mark)));
        return note;
    }

    /// <summary>That is how Word finds the separators; only in a new part.</summary>
    private static void DeclareSeparators(MainDocumentPart part, bool endnote, HashSet<string> touched)
    {
        var settingsPart = part.DocumentSettingsPart ?? part.AddNewPart<DocumentSettingsPart>();
        var settings = settingsPart.Settings ??= new Settings();

        if (endnote ? settings.GetFirstChild<EndnoteDocumentWideProperties>() is not null
                    : settings.GetFirstChild<FootnoteDocumentWideProperties>() is not null)
        {
            return;
        }

        OpenXmlElement properties = endnote
            ? new EndnoteDocumentWideProperties(
                new EndnoteSpecialReference { Id = -1 },
                new EndnoteSpecialReference { Id = 0 })
            : new FootnoteDocumentWideProperties(
                new FootnoteSpecialReference { Id = -1 },
                new FootnoteSpecialReference { Id = 0 });
        if (!settings.AddChild(properties, throwOnError: false)) return;

        settings.Save();
        touched.Add(settingsPart.Uri.ToString().TrimStart('/'));
    }

    /// <summary>
    /// Only when it differs from the package. Goes to <c>settings.xml</c> and to the last
    /// <c>w:sectPr</c>, if it also declares it, because that one wins. Absence does not erase.
    /// </summary>
    public static void ApplyNumbering(MainDocumentPart part, NotesDto? wanted, HashSet<string> touched)
    {
        if (wanted is null) return;
        var body = part.Document?.Body;
        if (body is null) return;
        var current = NotesReader.Read(part, body);
        var footnote = Differs(wanted.FootnotePr, current?.FootnotePr);
        var endnote = Differs(wanted.EndnotePr, current?.EndnotePr);
        if (!footnote && !endnote) return;

        var settingsPart = part.DocumentSettingsPart ?? part.AddNewPart<DocumentSettingsPart>();
        var settings = settingsPart.Settings ??= new Settings();
        var section = body.Elements<SectionProperties>().LastOrDefault();

        if (footnote)
        {
            var properties = settings.GetFirstChild<FootnoteDocumentWideProperties>();
            if (properties is null)
            {
                properties = new FootnoteDocumentWideProperties();
                foreach (var id in SpecialIds(part.FootnotesPart?.Footnotes))
                    properties.AppendChild(new FootnoteSpecialReference { Id = id });
                settings.AddChild(properties, throwOnError: false);
            }

            SetNumbering(properties, wanted.FootnotePr, () => new FootnotePosition());
            if (section?.GetFirstChild<FootnoteProperties>() is { } own)
                SetNumbering(own, wanted.FootnotePr, () => new FootnotePosition());
        }

        if (endnote)
        {
            var properties = settings.GetFirstChild<EndnoteDocumentWideProperties>();
            if (properties is null)
            {
                properties = new EndnoteDocumentWideProperties();
                foreach (var id in SpecialIds(part.EndnotesPart?.Endnotes))
                    properties.AppendChild(new EndnoteSpecialReference { Id = id });
                settings.AddChild(properties, throwOnError: false);
            }

            SetNumbering(properties, wanted.EndnotePr, () => new EndnotePosition());
            if (section?.GetFirstChild<EndnoteProperties>() is { } own)
                SetNumbering(own, wanted.EndnotePr, () => new EndnotePosition());
        }

        settings.Save();
        touched.Add(settingsPart.Uri.ToString().TrimStart('/'));
    }

    /// <summary>Absent means Word's.</summary>
    private static bool Differs(NotePrDto? wanted, NotePrDto? current) =>
        (wanted ?? new NotePrDto()) != (current ?? new NotePrDto());

    /// <c>-1</c> and <c>0</c> in Word.
    private static IEnumerable<long> SpecialIds(OpenXmlElement? notes) =>
        notes?.ChildElements.OfType<FootnoteEndnoteType>()
            .Where(note => note.Type?.Value is { } type &&
                           (type == FootnoteEndnoteValues.Separator || type == FootnoteEndnoteValues.ContinuationSeparator))
            .Select(note => note.Id?.Value ?? 0)
            .ToList() ?? [];

    /// <summary>In schema order, before the separator references.</summary>
    private static void SetNumbering(OpenXmlCompositeElement properties, NotePrDto? wanted, Func<OpenXmlElement> position)
    {
        string[] names = ["pos", "numFmt", "numStart", "numRestart"];
        foreach (var child in properties.ChildElements.Where(child => names.Contains(child.LocalName)).ToList())
            child.Remove();

        var values = new (OpenXmlElement Element, string? Value)[]
        {
            (position(), wanted?.Pos),
            (new NumberingFormat(), wanted?.NumFmt),
            (new NumberingStart(), wanted?.Start?.ToString(CultureInfo.InvariantCulture)),
            (new NumberingRestart(), wanted?.Restart),
        };
        OpenXmlElement? previous = null;
        foreach (var (element, value) in values)
        {
            if (value is null) continue;
            element.SetAttribute(new OpenXmlAttribute("w", "val", W, value));
            if (previous is null) properties.PrependChild(element);
            else previous.InsertAfterSelf(element);
            previous = element;
        }
    }

    /// <summary>A rewritten paragraph copies their <c>w:rPr</c>.</summary>
    public static Dictionary<string, Run> ReferenceRunsOf(MainDocumentPart part)
    {
        var runs = new Dictionary<string, Run>(StringComparer.Ordinal);
        foreach (var run in part.Document?.Body?.Descendants<Run>() ?? [])
        {
            if (BodyReader.NoteReferenceOf(run) is not FootnoteEndnoteReferenceType reference) continue;
            if (reference.Id?.Value is not { } id) continue;
            runs.TryAdd(Address(reference is EndnoteReference, id.ToString(CultureInfo.InvariantCulture)), run);
        }

        return runs;
    }
}

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;
using Cex = DocumentFormat.OpenXml.Office2021.Word.CommentsExt;

namespace Librevia.Format.Docx;

/// <summary>
/// Comparado comentário a comentário: o que não mudou fica com o XML de origem.
/// <list type="bullet">
/// <item>o editado é reescrito, com a formatação do primeiro parágrafo e o <c>w14:paraId</c>;</item>
/// <item>o novo recebe o id do modelo (outro, se já usado) e um <c>w14:paraId</c> inédito;</item>
/// <item>o que ficou sem âncora sai, com as respostas;</item>
/// <item>resolvido e resposta vão a <c>commentsExtended.xml</c>; autor novo, a <c>people.xml</c>.</item>
/// </list>
/// Antes de <c>MendCommentAnchors</c>, que precisa conhecer os comentários novos.
/// </summary>
public static class CommentsWriter
{
    /// <summary>O painel só edita texto simples.</summary>
    public const string RichEdited = "formatação de um comentário que você editou";

    private const string W14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private sealed class Entry(CommentDto model, Comment? original, string id)
    {
        public CommentDto Model { get; } = model;
        public Comment? Original { get; } = original;
        public string Id { get; } = id;
        public string? ParaId { get; set; }
    }

    public static void Apply(
        MainDocumentPart part, DocumentModelDto model, Inventory inventory, HashSet<string> touched)
    {
        // Rascunho anterior aos comentários, ou modelo que não fala deles: nada a fazer.
        if (model.BeforeComments || model.Comments is not { } wanted) return;
        var document = part.Document;
        if (document is null) return;

        var originals = new Dictionary<string, Comment>(StringComparer.Ordinal);
        foreach (var comment in part.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? [])
        {
            if (comment.Id?.Value is { Length: > 0 } id) originals.TryAdd(id, comment);
        }

        if (originals.Count == 0 && wanted.Count == 0) return;

        var before = Snapshot(part);
        var threads = CommentsReader.ThreadsOf(part);
        var paraIds = ParaIdsOf(part);
        var ids = new HashSet<string>(originals.Keys, StringComparer.Ordinal);
        ids.UnionWith(wanted.Select(comment => comment.Id));

        // Pelo id; o mesmo id com outro `paraId` é outro comentário, e ganha número novo.
        var entries = new List<Entry>();
        var byModelId = new Dictionary<string, Entry>(StringComparer.Ordinal);
        foreach (var wish in wanted)
        {
            if (byModelId.ContainsKey(wish.Id)) continue;
            var original = originals.GetValueOrDefault(wish.Id);
            var id = wish.Id;
            if (original is not null && !string.Equals(LastParaId(original), wish.ParaId, StringComparison.OrdinalIgnoreCase))
            {
                id = NextId(ids);
                Rename(AnchorRoots(part), wish.Id, id);
                original = null;
            }

            var entry = new Entry(wish, original, id);
            entry.ParaId = original is not null
                ? LastParaId(original)
                : wish.ParaId is { } carried && IsParaId(carried) && paraIds.Add(carried) ? carried : NewParaId(paraIds);
            entries.Add(entry);
            byModelId[wish.Id] = entry;
        }

        // A resposta nova não tem nó: as pontas vão ao lado das do comentário, como o Word grava.
        foreach (var entry in entries.Where(entry => entry.Original is null && entry.Model.ParentId is not null))
        {
            var rootId = byModelId.TryGetValue(entry.Model.ParentId!, out var root) ? root.Id : entry.Model.ParentId!;
            AddReplyAnchors(AnchorRoots(part), rootId, entry.Id);
        }

        // Fica o comentário com âncora em alguma parte, e a resposta cujo comentário fica.
        var anchored = AnchoredIds(part);
        var kept = entries.Where(entry => anchored.Contains(entry.Id)).ToList();
        var keptIds = kept.Select(entry => entry.Id).ToHashSet(StringComparer.Ordinal);
        kept.RemoveAll(entry =>
            entry.Model.ParentId is { } parent &&
            byModelId.TryGetValue(parent, out var root) &&
            !keptIds.Contains(root.Id));
        keptIds = kept.Select(entry => entry.Id).ToHashSet(StringComparer.Ordinal);

        var keptOriginals = kept.Where(entry => entry.Original is not null).Select(entry => entry.Original!).ToHashSet();
        var pruned = originals.Values
            .Where(comment => !keptOriginals.Contains(comment) &&
                              (!anchored.Contains(comment.Id!.Value!) || byModelId.ContainsKey(comment.Id!.Value!)))
            .ToHashSet();
        foreach (var comment in originals.Values)
        {
            if (LastParaId(comment) is { } paraId &&
                threads.GetValueOrDefault(paraId)?.Parent is { } parentId &&
                originals.GetValueOrDefault(parentId) is { } parent &&
                pruned.Contains(parent))
            {
                pruned.Add(comment);
            }
        }

        foreach (var entry in kept.Where(entry => entry.Original is not null && pruned.Contains(entry.Original)).ToList())
        {
            kept.Remove(entry);
        }

        var prunedParaIds = pruned.Select(LastParaId).OfType<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);

        var news = kept.Where(entry => entry.Original is null).ToList();
        var extended = new List<(string ParaId, bool Done, string? Parent)>();
        var noted = false;
        var edited = false;
        foreach (var entry in kept.Where(entry => entry.Original is not null))
        {
            var original = entry.Original!;
            var texts = original.Elements<Paragraph>().Select(CommentsReader.TextOf).ToList();
            if (!texts.SequenceEqual(entry.Model.Paragraphs, StringComparer.Ordinal))
            {
                if (entry.Model.Rich && !noted)
                {
                    inventory.NoteLoss(RichEdited);
                    noted = true;
                }

                edited = true;
                var properties = original.Elements<Paragraph>().FirstOrDefault()?.ParagraphProperties;
                original.RemoveAllChildren();
                foreach (var paragraph in ParagraphsOf(entry.Model.Paragraphs, properties, entry.ParaId))
                {
                    original.AppendChild(paragraph);
                }
            }

            var done = entry.ParaId is { } paraId && threads.GetValueOrDefault(paraId)?.Done == true;
            if (done == entry.Model.Done) continue;
            if (entry.ParaId is null)
            {
                // O de antes do Word 2013 não tem `paraId`: ganha um para o resolvido.
                entry.ParaId = NewParaId(paraIds);
                original.Elements<Paragraph>().LastOrDefault()?.SetAttribute(
                    new OpenXmlAttribute("w14", "paraId", W14Ns, entry.ParaId));
                DeclareW14(original);
            }

            extended.Add((entry.ParaId, entry.Model.Done, null));
        }

        if (news.Count > 0 || pruned.Count > 0 || edited)
        {
            var commentsPart = part.WordprocessingCommentsPart ?? part.AddNewPart<WordprocessingCommentsPart>();
            var root = commentsPart.Comments ??= NewCommentsRoot();
            foreach (var comment in pruned) comment.Remove();

            foreach (var entry in news)
            {
                // Sem o original (rascunho num pacote novo), o formatado sai com o texto simples.
                if (entry.Model.Rich && !noted)
                {
                    inventory.NoteLoss(RichEdited);
                    noted = true;
                }

                var comment = new Comment
                {
                    Id = entry.Id,
                    Author = entry.Model.Author,
                    Initials = entry.Model.Initials is { Length: > 0 } initials ? initials : InitialsOf(entry.Model.Author),
                };
                if (DateTime.TryParse(entry.Model.Date, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind,
                        out var date))
                {
                    comment.Date = date;
                }

                foreach (var paragraph in ParagraphsOf(entry.Model.Paragraphs, null, entry.ParaId))
                {
                    comment.AppendChild(paragraph);
                }

                root.AppendChild(comment);
                DeclareW14(root);

                var parent = entry.Model.ParentId is { } parentId && byModelId.TryGetValue(parentId, out var owner)
                    ? EnsureParaId(owner, paraIds, extended)
                    : null;
                extended.Add((entry.ParaId!, entry.Model.Done, parent));
            }
        }

        WriteExtended(part, extended, prunedParaIds);
        var removedDurable = WriteIds(part, news, prunedParaIds, out var durable);
        WriteExtensible(part, news, durable, removedDurable);
        WritePeople(part, news);

        foreach (var owned in CommentParts(part))
        {
            var now = owned.RootElement?.OuterXml;
            if (now is null) continue;
            var path = owned.Uri.OriginalString.TrimStart('/');
            if (before.TryGetValue(path, out var then) && then == now) continue;
            owned.RootElement!.Save();
            touched.Add(path);
        }
    }

    /// <summary>A resposta aponta o <c>paraId</c> do pai.</summary>
    private static string? EnsureParaId(
        Entry owner, HashSet<string> paraIds, List<(string ParaId, bool Done, string? Parent)> extended)
    {
        if (owner.ParaId is not null) return owner.ParaId;
        var last = owner.Original?.Elements<Paragraph>().LastOrDefault();
        if (last is null) return null;
        owner.ParaId = NewParaId(paraIds);
        last.SetAttribute(new OpenXmlAttribute("w14", "paraId", W14Ns, owner.ParaId));
        DeclareW14(owner.Original!);
        extended.Add((owner.ParaId, owner.Model.Done, null));
        return owner.ParaId;
    }

    private static Comments NewCommentsRoot()
    {
        var root = new Comments();
        root.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        root.AddNamespaceDeclaration("mc", McNs);
        root.AddNamespaceDeclaration("w14", W14Ns);
        root.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14" };
        return root;
    }

    /// <summary>No elemento de cima, para não repetir o <c>xmlns</c> em cada parágrafo.</summary>
    private static void DeclareW14(OpenXmlElement element)
    {
        var root = element.Ancestors<Comments>().FirstOrDefault() ?? element as Comments;
        if (root is not null && root.LookupNamespace("w14") is null) root.AddNamespaceDeclaration("w14", W14Ns);
    }

    /// <summary>O primeiro leva o <c>w:annotationRef</c>, e o último o <c>w14:paraId</c>.</summary>
    private static List<Paragraph> ParagraphsOf(
        IReadOnlyList<string> lines, ParagraphProperties? properties, string? paraId)
    {
        var texts = lines.Count == 0 ? [string.Empty] : lines;
        var result = new List<Paragraph>();
        for (var index = 0; index < texts.Count; index++)
        {
            var paragraph = new Paragraph();
            if (properties is not null) paragraph.AppendChild((ParagraphProperties)properties.CloneNode(true));
            if (index == 0)
            {
                paragraph.AppendChild(new Run(
                    new RunProperties(new RunStyle { Val = "CommentReference" }),
                    new AnnotationReferenceMark()));
            }

            if (texts[index].Length > 0) paragraph.AppendChild(RunOf(texts[index]));
            result.Add(paragraph);
        }

        if (paraId is not null) result[^1].ParagraphId = paraId;
        return result;
    }

    private static Run RunOf(string line)
    {
        var run = new Run();
        var piece = new System.Text.StringBuilder();
        void Flush()
        {
            if (piece.Length == 0) return;
            run.AppendChild(new Text(piece.ToString()) { Space = SpaceProcessingModeValues.Preserve });
            piece.Clear();
        }

        foreach (var character in line)
        {
            switch (character)
            {
                case '\t': Flush(); run.AppendChild(new TabChar()); break;
                case '\n': Flush(); run.AppendChild(new Break()); break;
                default: piece.Append(character); break;
            }
        }

        Flush();
        return run;
    }

    /// <summary>Começo com começo, fim com fim.</summary>
    private static void AddReplyAnchors(IReadOnlyList<OpenXmlElement> roots, string rootId, string replyId)
    {
        var document = roots.FirstOrDefault(root =>
            root.Descendants<CommentRangeStart>().Any(element => element.Id?.Value == rootId) ||
            root.Descendants<CommentReference>().Any(element => element.Id?.Value == rootId));
        if (document is null) return;
        var start = document.Descendants<CommentRangeStart>().FirstOrDefault(element => element.Id?.Value == rootId);
        if (start is not null)
        {
            OpenXmlElement after = start;
            while (after.NextSibling() is CommentRangeStart next) after = next;
            after.InsertAfterSelf(new CommentRangeStart { Id = replyId });
        }

        var reference = document.Descendants<CommentReference>().FirstOrDefault(element => element.Id?.Value == rootId);
        if (reference is null) return;
        OpenXmlElement anchor = reference.Parent is Run run && BodyReader.ReferenceOnly(run) is not null
            ? run
            : reference;
        while (anchor.NextSibling() is { } next &&
               (next is CommentRangeEnd || next is Run sibling && BodyReader.ReferenceOnly(sibling) is not null))
        {
            anchor = next;
        }

        if (start is not null)
        {
            var end = new CommentRangeEnd { Id = replyId };
            anchor.InsertAfterSelf(end);
            anchor = end;
        }

        anchor.InsertAfterSelf(new Run(
            new RunProperties(new RunStyle { Val = "CommentReference" }),
            new CommentReference { Id = replyId }));
    }

    private static HashSet<string> AnchoredIds(MainDocumentPart part)
    {
        var roots = new List<OpenXmlElement?> { part.Document };
        roots.AddRange(part.HeaderParts.Select(header => (OpenXmlElement?)header.Header));
        roots.AddRange(part.FooterParts.Select(footer => (OpenXmlElement?)footer.Footer));
        roots.Add(part.FootnotesPart?.Footnotes);
        roots.Add(part.EndnotesPart?.Endnotes);

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var root in roots.OfType<OpenXmlElement>())
        {
            foreach (var element in root.Descendants())
            {
                var id = element switch
                {
                    CommentRangeStart start => start.Id?.Value,
                    CommentRangeEnd end => end.Id?.Value,
                    CommentReference reference => reference.Id?.Value,
                    _ => null,
                };
                if (id is not null) ids.Add(id);
            }
        }

        return ids;
    }

    /// <summary>O corpo e as notas; as faixas o editor não tem.</summary>
    internal static IReadOnlyList<OpenXmlElement> AnchorRoots(MainDocumentPart part) =>
        new OpenXmlElement?[] { part.Document, part.FootnotesPart?.Footnotes, part.EndnotesPart?.Endnotes }
            .OfType<OpenXmlElement>().ToList();

    private static void Rename(IReadOnlyList<OpenXmlElement> roots, string from, string to)
    {
        foreach (var element in roots.SelectMany(root => root.Descendants()))
        {
            switch (element)
            {
                case CommentRangeStart start when start.Id?.Value == from: start.Id = to; break;
                case CommentRangeEnd end when end.Id?.Value == from: end.Id = to; break;
                case CommentReference reference when reference.Id?.Value == from: reference.Id = to; break;
            }
        }
    }

    /// <summary><c>w15:done</c> e <c>w15:paraIdParent</c>.</summary>
    private static void WriteExtended(
        MainDocumentPart part, List<(string ParaId, bool Done, string? Parent)> changes, HashSet<string> pruned)
    {
        var existing = part.WordprocessingCommentsExPart;
        if (changes.Count == 0 && (existing is null || pruned.Count == 0)) return;

        var owner = existing ?? part.AddNewPart<WordprocessingCommentsExPart>();
        var root = owner.CommentsEx ??= Ignorable(new W15.CommentsEx(), "w15");
        foreach (var entry in root.Elements<W15.CommentEx>().ToList())
        {
            if (entry.ParaId?.Value is { } paraId && pruned.Contains(paraId)) entry.Remove();
        }

        foreach (var (paraId, done, parent) in changes)
        {
            var entry = root.Elements<W15.CommentEx>().FirstOrDefault(candidate =>
                string.Equals(candidate.ParaId?.Value, paraId, StringComparison.OrdinalIgnoreCase));
            if (entry is null)
            {
                entry = new W15.CommentEx { ParaId = paraId };
                if (parent is not null) entry.ParaIdParent = parent;
                root.AppendChild(entry);
            }

            // "1" e "0", como o Word grava.
            entry.Done = new OnOffValue { InnerText = done ? "1" : "0" };
        }
    }

    /// <summary>Só quando o pacote já tem a parte. Devolve os <c>durableId</c> que saíram.</summary>
    private static HashSet<string> WriteIds(
        MainDocumentPart part, List<Entry> news, HashSet<string> pruned, out Dictionary<string, string> durable)
    {
        durable = new Dictionary<string, string>(StringComparer.Ordinal);
        var removed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var root = part.WordprocessingCommentsIdsPart?.CommentsIds;
        if (root is null || (news.Count == 0 && pruned.Count == 0)) return removed;

        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in root.Elements<Cid.CommentId>().ToList())
        {
            if (entry.DurableId?.Value is { } id) used.Add(id);
            if (entry.ParaId?.Value is { } paraId && pruned.Contains(paraId))
            {
                if (entry.DurableId?.Value is { } gone) removed.Add(gone);
                entry.Remove();
            }
        }

        foreach (var entry in news)
        {
            var id = NewHex(used);
            durable[entry.Id] = id;
            root.AppendChild(new Cid.CommentId { ParaId = entry.ParaId, DurableId = id });
        }

        return removed;
    }

    /// <summary>Só quando o pacote já tem a parte.</summary>
    private static void WriteExtensible(
        MainDocumentPart part, List<Entry> news, Dictionary<string, string> durable, HashSet<string> removed)
    {
        var root = part.WordCommentsExtensiblePart?.CommentsExtensible;
        if (root is null || (durable.Count == 0 && removed.Count == 0)) return;

        foreach (var entry in root.Elements<Cex.CommentExtensible>().ToList())
        {
            if (entry.DurableId?.Value is { } id && removed.Contains(id)) entry.Remove();
        }

        foreach (var entry in news)
        {
            if (!durable.TryGetValue(entry.Id, out var id)) continue;
            var extensible = new Cex.CommentExtensible { DurableId = id };
            if (DateTime.TryParse(entry.Model.Date, CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var date))
            {
                extensible.DateUtc = date;
            }

            root.AppendChild(extensible);
        }
    }

    /// <summary>Como o Word registra.</summary>
    private static void WritePeople(MainDocumentPart part, List<Entry> news)
    {
        var authors = news.Select(entry => entry.Model.Author).Where(author => author.Length > 0)
            .Distinct(StringComparer.Ordinal).ToList();
        if (authors.Count == 0) return;

        var known = part.WordprocessingPeoplePart?.People?.Elements<W15.Person>()
            .Select(person => person.Author?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal) ?? [];
        var missing = authors.Where(author => !known.Contains(author)).ToList();
        if (missing.Count == 0) return;

        var owner = part.WordprocessingPeoplePart ?? part.AddNewPart<WordprocessingPeoplePart>();
        var root = owner.People ??= Ignorable(new W15.People(), "w15");
        foreach (var author in missing)
        {
            root.AppendChild(new W15.Person(new W15.PresenceInfo { ProviderId = "None", UserId = author })
            {
                Author = author,
            });
        }
    }

    private static T Ignorable<T>(T root, string prefix) where T : OpenXmlPartRootElement
    {
        root.AddNamespaceDeclaration("mc", McNs);
        root.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = prefix };
        return root;
    }

    private static IEnumerable<OpenXmlPart> CommentParts(MainDocumentPart part) =>
        new OpenXmlPart?[]
        {
            part.WordprocessingCommentsPart, part.WordprocessingCommentsExPart, part.WordprocessingCommentsIdsPart,
            part.WordCommentsExtensiblePart, part.WordprocessingPeoplePart,
        }.OfType<OpenXmlPart>();

    private static Dictionary<string, string> Snapshot(MainDocumentPart part) =>
        CommentParts(part)
            .Where(owned => owned.RootElement is not null)
            .ToDictionary(owned => owned.Uri.OriginalString.TrimStart('/'), owned => owned.RootElement!.OuterXml,
                StringComparer.Ordinal);

    private static string? LastParaId(Comment comment) =>
        comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;

    /// <summary>O novo não pode repetir nenhum.</summary>
    private static HashSet<string> ParaIdsOf(MainDocumentPart part)
    {
        var roots = new List<OpenXmlElement?> { part.Document, part.WordprocessingCommentsPart?.Comments };
        roots.AddRange(part.HeaderParts.Select(header => (OpenXmlElement?)header.Header));
        roots.AddRange(part.FooterParts.Select(footer => (OpenXmlElement?)footer.Footer));
        roots.Add(part.FootnotesPart?.Footnotes);
        roots.Add(part.EndnotesPart?.Endnotes);
        return roots.OfType<OpenXmlElement>()
            .SelectMany(root => root.Descendants<Paragraph>())
            .Select(paragraph => paragraph.ParagraphId?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Oito dígitos hexadecimais, abaixo de 0x80000000.</summary>
    private static bool IsParaId(string value) =>
        value.Length == 8 &&
        int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var number) &&
        number > 0;

    private static string NewParaId(HashSet<string> used) => NewHex(used);

    private static string NewHex(HashSet<string> used)
    {
        while (true)
        {
            var candidate = Random.Shared.Next(1, int.MaxValue).ToString("X8", CultureInfo.InvariantCulture);
            if (used.Add(candidate)) return candidate;
        }
    }

    private static string NextId(HashSet<string> used)
    {
        var next = used.Select(id => int.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out var n) ? n : -1)
            .DefaultIfEmpty(-1).Max() + 1;
        while (!used.Add(next.ToString(CultureInfo.InvariantCulture))) next++;
        return next.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Como o Word: a primeira letra de cada palavra.</summary>
    internal static string InitialsOf(string author) =>
        string.Concat(author.Split(' ', StringSplitOptions.RemoveEmptyEntries).Take(3)
            .Select(word => char.ToUpperInvariant(word[0])));
}

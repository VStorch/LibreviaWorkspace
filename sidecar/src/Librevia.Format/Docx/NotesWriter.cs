using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O corpo das notas de rodapé e de fim volta para `footnotes.xml` e `endnotes.xml` (M11).
/// </summary>
/// <remarks>
/// A mesma aposta do corpo e das faixas: a nota cujo corpo não mudou não é tocada,
/// e a parte em que nenhuma nota mudou nem entra na lista de graváveis — volta
/// byte a byte (ver DocxWriter.RestoreUntouchedParts). Na nota editada, só o
/// parágrafo editado é reescrito; os outros voltam como estavam.
///
/// O número da nota não mora em lugar nenhum do modelo: é a ordem da referência
/// no documento. O que mora é o `nid`, o `w:id` que casa a referência com o
/// corpo. A nota nova (`nid` nulo) e a repetida (a referência colada duas vezes)
/// ganham um id acima do maior da parte; a nota cuja referência sumiu do texto
/// sai da parte, como no Word.
/// </remarks>
internal static class NotesWriter
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public const string Footnote = "footnote";
    public const string Endnote = "endnote";

    /// <summary>O endereço de uma nota: `fn:3`, `en:1`. Os blocos dela são `fn:3/p1`…</summary>
    public static string Address(bool endnote, string id) => (endnote ? "en:" : "fn:") + id;

    /// <summary>Uma referência do modelo e a nota que ela vai apontar no arquivo.</summary>
    /// <param name="Fresh">A nota não existe no arquivo: será criada com <paramref name="Id"/>.</param>
    /// <param name="Original">O id que a nota tinha no arquivo, quando a renumeração o trocou.</param>
    public sealed record Wanted(Node Reference, bool Endnote, string Id, bool Fresh, string? Original = null)
    {
        /// <summary>O id pelo qual a nota é achada no arquivo original.</summary>
        public string Source => Original ?? Id;
    }

    /// <summary>A nota normal (não o separador) com este id, quando a parte a tem.</summary>
    public static OpenXmlElement? NoteOf(MainDocumentPart part, bool endnote, string id) =>
        NotesIn(part, endnote).FirstOrDefault(note => IsNormal(note) && IdOf(note) == id);

    private static IEnumerable<FootnoteEndnoteType> NotesIn(MainDocumentPart part, bool endnote) =>
        endnote
            ? part.EndnotesPart?.Endnotes?.Elements<DocumentFormat.OpenXml.Wordprocessing.Endnote>() ?? []
            : (IEnumerable<FootnoteEndnoteType>?)part.FootnotesPart?.Footnotes?.Elements<DocumentFormat.OpenXml.Wordprocessing.Footnote>() ?? [];

    /// <summary>Separador, separador de continuação e aviso de continuação não são notas do texto.</summary>
    private static bool IsNormal(FootnoteEndnoteType note) =>
        note.Type?.Value is null || note.Type.Value == FootnoteEndnoteValues.Normal;

    private static string? IdOf(FootnoteEndnoteType note) =>
        note.Id?.Value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Decide o id de cada referência do modelo, antes de o corpo ser gravado.
    /// </summary>
    /// <remarks>
    /// Antes porque o run da referência leva o id, e quem o grava é o parágrafo.
    /// A referência que ganha id novo muda de `nid` no modelo — e o parágrafo que
    /// a leva passa a ser reescrito, que é o certo: a referência dele mudou.
    /// </remarks>
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
    /// Ids na ordem do texto, como o Word grava — quando uma referência movida a
    /// tirou dela.
    /// </summary>
    /// <remarks>
    /// O LibreOffice ignora o `w:id` e dá às referências, na ordem do texto, as
    /// notas na ordem dos ids: a nota movida (recortar e colar) aparecia com o
    /// corpo de outra. O arquivo que já vinha fora de ordem e a gravação que não
    /// moveu nada ficam como estão — renumerar reescreveria os parágrafos à toa.
    /// </remarks>
    private static List<Wanted> InTextOrder(List<Wanted> wanted, MainDocumentPart part)
    {
        var result = new List<Wanted>(wanted);
        foreach (var endnote in new[] { false, true })
        {
            if (!Ascending(OriginalOrder(part, endnote))) continue;
            var mine = result.Select((entry, index) => (entry, index)).Where(pair => pair.entry.Endnote == endnote).ToList();
            // A nota nova também conta: inserida antes das outras, ela ganha o maior id
            // e quebra a ordem do mesmo jeito que uma movida.
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

    /// <summary>Os ids das notas na ordem em que o corpo do arquivo as referencia.</summary>
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

    /// <summary>Os `noteRef` do documento, em ordem — sem descer no corpo de nenhum.</summary>
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

    /// <summary>
    /// Grava o corpo das notas que mudaram, cria as novas e tira as que perderam a
    /// referência. Devolve quantos blocos de nota foram reescritos.
    /// </summary>
    /// <param name="read">As notas que a leitura de referência achou no corpo — ver BodyReader.Notes.</param>
    /// <param name="writerFor">O escritor de parágrafos dono dos relacionamentos de cada parte.</param>
    public static int Apply(
        MainDocumentPart part,
        List<Wanted> wanted,
        IReadOnlyDictionary<string, BodyReader.NoteRead> read,
        Func<OpenXmlPart, ParagraphWriter> writerFor,
        NumberingFactory numbering,
        Inventory inventory,
        HashSet<string> touched,
        bool beforeComments,
        bool beforeRevisions)
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

            // O LibreOffice casa a nota com a referência pela ordem em que as notas
            // aparecem na parte, e não pelo id: a nota movida no texto (recortar e
            // colar) tem de mudar de lugar na parte também, como o Word as grava.
            // Comparada com a ordem das referências no corpo original, e não com a
            // da parte: um arquivo que já vinha fora de ordem e não foi mexido volta
            // byte a byte.
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
                            beforeComments, beforeRevisions);
                        EnsureReferenceMark(original.Source, mark, endnote);
                    }

                    continue;
                }

                var note = endnote
                    ? (FootnoteEndnoteType)new DocumentFormat.OpenXml.Wordprocessing.Endnote()
                    : new DocumentFormat.OpenXml.Wordprocessing.Footnote();
                note.Id = long.Parse(entry.Id, CultureInfo.InvariantCulture);
                var empty = new BodyReader.NoteRead(note, []);
                rewritten += Rebuild(empty, entry.Reference, writer, numbering, inventory, beforeComments, beforeRevisions);
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

    /// <summary>`word/footnotes.xml` → `word/_rels/footnotes.xml.rels`.</summary>
    private static string RelationshipsPathOf(string path)
    {
        var slash = path.LastIndexOf('/');
        return slash < 0 ? $"_rels/{path}.rels" : $"{path[..slash]}/_rels/{path[(slash + 1)..]}.rels";
    }

    /// <summary>
    /// O corpo da nota no modelo é o do arquivo? Os mesmos blocos, na mesma ordem,
    /// cada um com a impressão digital que a leitura deu — a comparação do corpo.
    /// </summary>
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

    /// <summary>
    /// Refaz o corpo de uma nota: o parágrafo que não mudou volta como estava, o
    /// editado é reescrito. Devolve quantos blocos foram reescritos.
    /// </summary>
    private static int Rebuild(
        BodyReader.NoteRead original,
        Node reference,
        ParagraphWriter writer,
        NumberingFactory numbering,
        Inventory inventory,
        bool beforeComments,
        bool beforeRevisions)
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

            if (!DocxWriter.BuildSlot(slot, owner, writer, inventory, elements, beforeComments, beforeRevisions, beforeNotes: false))
            {
                rewritten++;
            }

            if (owner is not null)
            {
                foreach (var loose in owner.Trailing) elements.Add(loose.CloneNode(true));
            }
        }

        // O que a nota apagada levava entre os blocos — um marcador — vai com ela.
        if (elements.Count == 0) elements.Add(new Paragraph());

        var source = original.Source;
        source.RemoveAllChildren();
        foreach (var element in elements) source.AppendChild(element);
        return rewritten;
    }

    /// <summary>
    /// O número no começo da nota (`w:footnoteRef`): o primeiro parágrafo tem de
    /// trazê-lo, e o leitor não o deu ao modelo.
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

    /// <summary>O run do `w:footnoteRef` da parte, para copiar — ou um no estilo do Word.</summary>
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
    /// O `w:rPr` da referência que nasce no editor: o estilo de caractere do Word
    /// quando o documento o define, e o sobrescrito direto quando não — sem ele, o
    /// número sairia na linha do texto.
    /// </summary>
    public static RunProperties ReferenceProperties(MainDocumentPart part, bool endnote)
    {
        var style = endnote ? "EndnoteReference" : "FootnoteReference";
        return StyleDefined(part, style)
            ? new RunProperties(new RunStyle { Val = style })
            : new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
    }

    /// <summary>A nota nova com o estilo de texto de nota, quando o documento o tem e o parágrafo não traz outro.</summary>
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
    /// A parte das notas — criada, com os separadores que o Word exige, quando o
    /// documento ainda não tem nenhuma nota deste tipo.
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

    /// <summary>
    /// `w:footnotePr` no `settings.xml`, apontando os dois separadores — é assim
    /// que o Word os acha. Só quando a parte nasceu agora e o arquivo não os aponta.
    /// </summary>
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
    /// Grava a numeração das notas (`w:footnotePr`/`w:endnotePr`) quando a do modelo
    /// difere da que o pacote declara. Igual, nada se toca: o `settings.xml` do
    /// arquivo aberto e salvo volta byte a byte.
    /// </summary>
    /// <remarks>
    /// Vai para o `settings.xml`, que vale para o documento todo; e, se o último
    /// `w:sectPr` também declara a numeração — é ela que vence na leitura —, para
    /// ele também. Modelo sem numeração não apaga a do pacote: ausência não é pedido.
    /// </remarks>
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

    /// <summary>A numeração ausente é a do Word: só os campos declarados contam.</summary>
    private static bool Differs(NotePrDto? wanted, NotePrDto? current) =>
        (wanted ?? new NotePrDto()) != (current ?? new NotePrDto());

    /// <summary>Os ids das notas separadoras da parte (`-1` e `0`, no Word), que o `w:footnotePr` aponta.</summary>
    private static IEnumerable<long> SpecialIds(OpenXmlElement? notes) =>
        notes?.ChildElements.OfType<FootnoteEndnoteType>()
            .Where(note => note.Type?.Value is { } type &&
                           (type == FootnoteEndnoteValues.Separator || type == FootnoteEndnoteValues.ContinuationSeparator))
            .Select(note => note.Id?.Value ?? 0)
            .ToList() ?? [];

    /// <summary>
    /// Troca `w:pos`, `w:numFmt`, `w:numStart` e `w:numRestart` pelos do modelo, na
    /// ordem do esquema — antes das referências às separadoras.
    /// </summary>
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

    /// <summary>
    /// O run de cada referência de nota do corpo original, pelo endereço — é dele
    /// que o parágrafo reescrito copia o `w:rPr` (o estilo, o sobrescrito).
    /// </summary>
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

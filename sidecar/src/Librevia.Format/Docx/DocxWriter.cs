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
/// Gravação cirúrgica: reescreve só o que o usuário tocou.
/// </summary>
/// <remarks>
/// O desenho está em docs/02-docx-cirurgico.md. Em uma frase: a fidelidade não
/// vem de entender o OOXML, vem de **não mexer** no que não foi editado.
///
/// Só `word/document.xml` é reescrito. Estilos, numeração, cabeçalhos,
/// rodapés, mídia, comentários, notas, tema e configurações continuam
/// exatamente como estavam, a não ser que o usuário os tenha mudado.
/// </remarks>
public static class DocxWriter
{
    public static (byte[] Bytes, SaveResult Result) Write(byte[] original, DocumentModelDto model)
    {
        var inventory = new Inventory();

        // Trabalha numa cópia gravável: `original` são os bytes que o processo
        // main guardou na abertura e não podem ser alterados.
        using var buffer = new MemoryStream();
        buffer.Write(original, 0, original.Length);
        buffer.Position = 0;

        using var document = OpenEditable(buffer);
        var part = document.MainDocumentPart
                   ?? throw new DocxException("O arquivo original não contém um documento do Word.");
        var body = part.Document?.Body
                   ?? throw new DocxException("O arquivo original está vazio ou danificado.");

        // Reindexa a partir dos mesmos bytes: os ids saem iguais aos da
        // abertura porque a numeração é posicional e determinística.
        //
        // Com o mesmo leitor que produziu o modelo: um rascunho antigo traz os
        // blocos achatados, e comparados com a leitura que só leva o direto todo
        // bloco pareceria mudado — o documento inteiro seria reescrito.
        var reader = new BodyReader(
            part,
            new Inventory(),
            model.Flatten,
            !model.BeforeReferences,
            !model.BeforeSections,
            !model.BeforeComments,
            !model.BeforeRevisions,
            !model.BeforeNotes);
        var (_, blocks) = reader.Read(body);
        var index = blocks.ToDictionary(block => block.Oid, StringComparer.Ordinal);

        var section = body.Elements<SectionProperties>().LastOrDefault();

        // As partes fora de `word/document.xml` que esta gravação tem o direito
        // de mexer. Começa vazio e só cresce quando algo de fato muda — uma faixa
        // em que se digitou, a numeração de uma lista nova.
        var touched = new HashSet<string>(StringComparer.Ordinal);

        // Os estilos antes do corpo: o bloco que aponta um estilo criado agora
        // precisa encontrá-lo definido, e o resolvedor de quem grava os blocos
        // precisa ler as definições novas. No rascunho antigo, só o estilo que o
        // pacote não tem: os dele podem ter sido inventados na migração (ver
        // StyleWriter.Apply), e a mudança num existente vira perda declarada.
        StyleWriter.Apply(part, model.Styles, inventory, touched, additionsOnly: model.Flatten);

        // O id de cada referência de nota (M11), antes do corpo: o run da
        // referência o leva, e quem grava o run é o parágrafo. No rascunho de antes
        // das notas o modelo não tem referência nenhuma, e as partes ficam como
        // estão — a perda do parágrafo editado é declarada em NoteWhatWasInside.
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

        // O corpo das notas (M11): só a nota que mudou, e só a parte que a guarda.
        // Antes dos comentários, que procuram as âncoras também nas notas.
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
                model.BeforeRevisions);
        }

        // A numeração das notas (M11): só quando o modelo pede outra que a do
        // pacote — o `.sdoc` reaberto e gravado como `.docx` não a perde mais.
        NotesWriter.ApplyNumbering(part, model.Notes, touched);

        // O corpo dos comentários: o criado, o editado, o resolvido e o excluído —
        // ver CommentsWriter. Antes do conserto das pontas, que precisa conhecer os
        // comentários novos para não descartar as âncoras deles.
        CommentsWriter.Apply(part, model, inventory, touched);

        // As pontas de comentário que a edição desemparelhou — ver MendCommentAnchors.
        // No corpo e nas notas (M11): o comentário criado numa nota sai de lá com a
        // referência, e o excluído leva as pontas que a nota preservada ainda tinha.
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

        // As revisões (M10): a movimentação que a edição partiu vira exclusão e
        // inserção, e cada revisão sai com um `w:id` só dela.
        MendMoves(body);
        Revisions.MakeIdsUnique(body, part);
        Revisions.ApplyTracking(part, model.TrackChanges, touched, inventory);

        // `w:sectPr` fecha o corpo e carrega a configuração de página.
        body.AppendChild(section is null ? new SectionProperties() : section);

        // Só se a página tiver mudado. Regravar `w:pgSz` e `w:pgMar` em todo save
        // custava o papel de quem não usa A4 nem Carta: o modelo só conhece esses
        // dois, e o arredondamento voltava para o arquivo como se fosse escolha do
        // autor. Ver PageReader.Matches.
        var current = body.Elements<SectionProperties>().Last();
        if (!PageReader.Matches(current, model.Page)) ApplyPageSetup(current, model.Page);

        // As seções antes da última: as marcas que o corpo levou, com a
        // configuração que o modelo dá a cada uma. No rascunho de antes das
        // seções não há marca nenhuma no modelo, e os `w:sectPr` de parágrafo
        // voltam como estavam — ver SectionWriter.
        var aliases = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!model.BeforeSections)
        {
            SectionWriter.ApplyStart(current, model.Page);
            SectionWriter.ApplyColumns(current, model.Page);
            aliases = SectionWriter.Apply(part, breaks, current, model, inventory, touched);
        }

        // O cabeçalho e o rodapé de texto simples do documento novo — ver
        // PlainBandWriter. No documento que veio de fora a faixa manda, e isto
        // não faz nada.
        PlainBandWriter.Apply(part, current, model.Page, inventory, touched);

        // Formato e início do número de página, capa distinta e páginas pares.
        PageNumbering.Apply(part, current, model.Page, touched, inventory);

        // O texto digitado no cabeçalho e no rodapé, peça por peça. Só as
        // partes que de fato mudaram entram na lista de graváveis: o resto
        // continua saindo do arquivo original, byte a byte.
        touched.UnionWith(BandWriter.Apply(part, [.. model.Sections ?? [], model.Page], inventory, aliases));

        // As propriedades (M11): só a parte em que algum campo mudou — ver
        // DocumentProperties. `docProps/custom.xml` nunca é tocado.
        DocumentProperties.Apply(document, model.Properties, touched);

        part.Document!.Save();
        document.Dispose();

        return (RestoreUntouchedParts(original, buffer.ToArray(), touched),
            new SaveResult(inventory, preserved, rewritten));
    }

    /// <summary>
    /// Partes que a gravação tem o direito de alterar. Todo o resto volta a ser
    /// exatamente o que era.
    /// </summary>
    /// <remarks>
    /// Fixa porque não depende do documento. A parte de cabeçalho ou rodapé em
    /// que se digitou entra por fora, uma a uma: qual delas é depende do que a
    /// pessoa tocou, e liberar todas de antemão devolveria a reserialização do
    /// SDK a um cabeçalho que ninguém abriu.
    /// </remarks>
    private static readonly HashSet<string> Writable = new(StringComparer.Ordinal)
    {
        "word/document.xml",
        // Mudam quando o usuário insere imagem ou link num bloco editado.
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
    };

    /// <summary>
    /// Devolve às demais partes o conteúdo original, byte a byte.
    /// </summary>
    /// <remarks>
    /// Não é zelo excessivo: o SDK **reserializa toda parte cujo DOM tipado foi
    /// materializado**, mesmo sem alteração nenhuma. Basta o leitor de
    /// numeração tocar em `NumberingDefinitionsPart.Numbering` para
    /// `word/numbering.xml` sair diferente de um documento que ninguém editou —
    /// medido no corpus real.
    ///
    /// Reescrita sem intenção é a porta por onde a fidelidade escapa, e ela
    /// escapa em silêncio. Impor a invariante aqui é mais seguro do que confiar
    /// em ninguém nunca materializar um DOM sem querer.
    /// </remarks>
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
    /// Deixa cada comentário do corpo com pontas que o Word e o LibreOffice leem.
    /// </summary>
    /// <remarks>
    /// O editor apaga uma ponta junto com o texto, e o parágrafo preservado guarda
    /// a dele: sobra começo sem fim, ou âncora de um comentário que o pacote não
    /// tem (o rascunho reaberto e gravado num pacote novo). O LibreOffice recusa o
    /// arquivo no segundo caso e descarta a conversa no primeiro. Então:
    /// <list type="bullet">
    /// <item>âncora de comentário que o pacote não tem sai — a perda já foi declarada na leitura;</item>
    /// <item>começo sem fim nem referência vira comentário de ponto ali mesmo;</item>
    /// <item>fim sem começo sai, e fica a referência;</item>
    /// <item>trecho inteiro sem referência ganha uma logo depois do fim.</item>
    /// </list>
    /// Só elementos de largura zero: o parágrafo preservado continua preservado.
    /// </remarks>
    /// <returns>Se mudou alguma coisa — a parte das notas só é gravada então.</returns>
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
            // A referência mora num run: sozinha nele, o run vai junto.
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
    /// A movimentação que continua inteira no corpo, e a que a edição partiu.
    /// </summary>
    /// <remarks>
    /// Mover é um par: o trecho de origem (`w:moveFrom`) e o de destino (`w:moveTo`),
    /// cada um entre as pontas de um intervalo com o mesmo nome. O parágrafo
    /// reescrito leva o trecho como `w:del`/`w:ins` e perde as pontas; o que sobra
    /// do par no parágrafo preservado viraria movimentação sem origem ou sem
    /// destino. Então o par incompleto vira exclusão e inserção dos dois lados, e
    /// as pontas soltas saem — a perda já foi declarada em NoteWhatWasInside.
    /// </remarks>
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

        // O nome com as duas pontas dos dois lados.
        var fromEndIds = fromEnds.Select(Id).ToHashSet(StringComparer.Ordinal);
        var toEndIds = toEnds.Select(Id).ToHashSet(StringComparer.Ordinal);
        var fromNames = fromStarts.Where(start => fromEndIds.Contains(Id(start))).Select(Name)
            .ToHashSet(StringComparer.Ordinal);
        var toNames = toStarts.Where(start => toEndIds.Contains(Id(start))).Select(Name)
            .ToHashSet(StringComparer.Ordinal);
        var whole = fromNames.Intersect(toNames, StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);

        // Os ids dos intervalos que ficam, e o lado de cada um.
        var keptIds = new HashSet<string?>(StringComparer.Ordinal);
        foreach (var start in fromStarts.Concat<OpenXmlElement>(toStarts))
        {
            if (whole.Contains(Name(start))) keptIds.Add((start is MoveFromRangeStart ? "f" : "t") + Id(start));
        }

        // Cada trecho de movimentação dentro de um intervalo que fica, em ordem.
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

            // Um `oid` repetido é bloco colado: o XML original é de **um** deles,
            // e preservá-lo duas vezes duplicaria âncoras de comentário, ids de
            // revisão e marcadores. Da segunda ocorrência em diante o bloco é
            // tratado como novo — e é por isso que a resposta fica guardada:
            // vale também para o que se copia do original mais abaixo.
            var first = oid is not null && used.Add(oid);

            // O que o corpo guardava entre este bloco e o anterior — um marcador
            // solto entre dois parágrafos, um controle de conteúdo que o leitor
            // não conhece — volta antes dele, preservado ou reescrito. Ver
            // Block.Leading. Só na primeira ocorrência, pela mesma razão do XML
            // do próprio bloco.
            var owner = first && index.TryGetValue(oid!, out var known) ? known : null;
            if (owner is not null)
            {
                foreach (var loose in owner.Leading) elements.Add(loose.CloneNode(true));
            }

            var before = elements.Count;
            var kept = BuildSlot(
                slot, owner, writer, inventory, elements, model.BeforeComments, model.BeforeRevisions, model.BeforeNotes);
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

        // O bloco apagado leva junto o que estava solto antes dele: um marcador
        // perdido com o texto que marcava é o que o Word também faz. Qualquer outra
        // coisa — um controle de conteúdo — não se apaga em silêncio.
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
    /// Um marcador por nome e por id, no documento inteiro.
    /// </summary>
    /// <remarks>
    /// O parágrafo copiado e colado leva os marcadores junto — são nós do
    /// modelo —, e o Word recusa dois `w:bookmarkStart` de mesmo id: é âncora
    /// ambígua para o sumário e para a referência que o cita. Quem cede é sempre
    /// o bloco gerado agora; o preservado volta como estava. Entre dois gerados,
    /// fica o primeiro, que é a ordem em que o Word resolve o nome repetido.
    /// </remarks>
    /// <summary>
    /// A ponta solta que ia embora com o bloco apagado, quando a outra ponta ficou.
    /// </summary>
    /// <remarks>
    /// O Word grava entre dois parágrafos o fim do marcador que termina depois de
    /// uma tabela. Apagado o bloco que o guardava, o começo ficava sem fim — e o
    /// marcador órfão é âncora quebrada para quem o cita. O fim volta logo depois
    /// do bloco do começo, e o começo órfão de fim, logo antes do bloco do fim: o
    /// marcador encolhe até o que sobrou dele, como no Word.
    /// </remarks>
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

    private static void UniqueBookmarks(List<OpenXmlElement> elements, HashSet<OpenXmlElement> generated)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ends = new HashSet<string>(StringComparer.Ordinal);

        // O maior id do corpo inteiro — inclusive os que o modelo não conhece: as
        // pontas entre linhas de tabela, as soltas entre blocos, as de dentro de
        // caixa de texto. É dele que sai o id novo de quem precisa renumerar.
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
        // Id antigo → id novo, do marcador renumerado que ainda espera a ponta final.
        var renamed = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var element in elements.Where(generated.Contains))
        {
            // As pontas na ordem do documento, para que o fim encontre o começo
            // renumerado que veio antes dele.
            foreach (var mark in element.Descendants().Prepend(element).ToList())
            {
                if (mark is BookmarkStart start)
                {
                    var id = start.Id?.Value ?? string.Empty;
                    var name = start.Name?.Value ?? string.Empty;

                    // O nome repetido é a cópia colada de um marcador que o
                    // documento já tem: sai, com a ponta final.
                    if (!names.Add(name))
                    {
                        dropped.Add(id);
                        start.Remove();
                        continue;
                    }

                    // Só o id repetido é outro marcador que calhou de ter o mesmo
                    // número — o de outro documento, que o Word também numera de
                    // zero. Ganha um id novo em vez de sumir.
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

    /// <summary>
    /// Grava um bloco: o XML original quando o conteúdo não mudou, o reescrito
    /// quando mudou. Devolve se foi preservado.
    /// </summary>
    /// <param name="owner">O bloco do arquivo com o mesmo `oid`, na primeira ocorrência dele.</param>
    internal static bool BuildSlot(
        Slot slot,
        Block? owner,
        ParagraphWriter writer,
        Inventory inventory,
        List<OpenXmlElement> elements,
        bool beforeComments,
        bool beforeRevisions,
        bool beforeNotes)
    {
        if (owner is not null && SameContent(slot, owner))
        {
            // O conteúdo é o mesmo, mas o item pode ter mudado de lugar na
            // lista: Tab o desce um nível, "Reiniciar numeração" o põe noutro
            // `w:num`. Nada disso está no item — está na lista em volta —, e
            // devolver o XML como estava desfazia a mudança ao reabrir. Só o
            // `w:numPr` é trocado; o resto do parágrafo volta como veio.
            if (slot.List is { } list && owner.Source is Paragraph paragraph && !Points(paragraph, list))
            {
                elements.Add(Renumbered(paragraph, list));
                return false;
            }

            elements.Add(owner.Source.CloneNode(true));
            foreach (var next in owner.Continuation) elements.Add(next.CloneNode(true));
            return true;
        }

        // O XML original do bloco editado ainda serve para o que este
        // escritor não sabe gerar: os objetos ancorados seguem para o
        // parágrafo reescrito em vez de sumirem com ele.
        var source = owner?.Source;

        // O embrulho é a afirmação do corpo: aqui se sabe se o parágrafo é
        // item de lista ou não. Quem grava dentro de uma célula não sabe, e
        // passa `null` — ver ParagraphWriter.ListPlacement.
        var placement = new ParagraphWriter.ListPlacement(slot.List);

        foreach (var element in writer.Write(slot.Content, placement, source)) elements.Add(element);

        if (owner is not null) NoteWhatWasInside(owner, inventory, beforeComments, beforeRevisions, beforeNotes);
        return false;
    }

    /// <summary>O bloco do modelo diz o mesmo que o do arquivo — ver OwnContent.</summary>
    private static bool SameContent(Slot slot, Block owner) =>
        string.Equals(
            OwnContent(owner.Extracted).Fingerprint(),
            OwnContent(slot.Identity).Fingerprint(),
            StringComparison.Ordinal);

    /// <summary>
    /// O bloco voltaria ao arquivo exatamente como estava: o mesmo conteúdo, e o
    /// item de lista na mesma numeração — ver BuildSlot.
    /// </summary>
    internal static bool Preservable(Slot slot, Block owner) =>
        SameContent(slot, owner) &&
        !(slot.List is { } list && owner.Source is Paragraph paragraph && !Points(paragraph, list));

    /// <summary>
    /// O item de lista sem as sublistas de dentro — o que corresponde ao `w:p` dele.
    /// </summary>
    /// <remarks>
    /// No editor a sublista mora dentro do item de cima; no arquivo ela são os
    /// parágrafos seguintes, cada um com a sua identidade. Comparada com elas, a
    /// impressão digital do item de cima mudava a cada Tab num item de baixo — e
    /// o parágrafo de cima, que ninguém tocou, era reescrito.
    /// </remarks>
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

    /// <summary>O parágrafo já aponta esta numeração, neste nível?</summary>
    private static bool Points(Paragraph paragraph, ParagraphWriter.ListContext list)
    {
        var numbering = paragraph.ParagraphProperties?.NumberingProperties;
        return numbering?.NumberingId?.Val?.Value == list.NumberingId &&
               (numbering.NumberingLevelReference?.Val?.Value ?? 0) == list.Level;
    }

    /// <summary>O parágrafo original, apontando a numeração e o nível novos.</summary>
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
    /// Achata a árvore do editor na sequência de blocos que o corpo do DOCX
    /// espera, desembrulhando listas.
    /// </summary>
    /// <remarks>
    /// No editor uma lista é um nó com itens dentro; no OOXML são parágrafos
    /// irmãos, cada um apontando a mesma numeração. É por isso que o `oid` mora
    /// no `listItem`: ele é que corresponde a um `w:p`.
    /// </remarks>
    /// <param name="Identity">
    /// O nó extraído na abertura — é contra ele que a impressão digital é
    /// comparada. Para um item de lista é o `listItem`, não o parágrafo de
    /// dentro: comparar coisas de tipos diferentes nunca daria igual, e **nada
    /// seria preservado**, transformando a edição cirúrgica em regeneração
    /// completa sem que nada falhasse visivelmente.
    /// </param>
    /// <param name="Content">O nó a gravar quando não houver preservação.</param>
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
                    // O nível que o arquivo deu à lista, quando a árvore não o
                    // diz sozinha; senão, um abaixo da lista de fora.
                    var level = Math.Clamp(
                        Attr.Int(node, "level") ?? (inherited?.Level ?? -1) + 1, 0, ListLevels.Count - 1);
                    var definition = Attr.Node(node, "numbering") as System.Text.Json.Nodes.JsonObject;

                    // A numeração da lista de fora só serve à de dentro quando as
                    // duas são do mesmo tipo: herdada às cegas, uma sublista
                    // numerada dentro de uma com marcador saía com marcador. E só
                    // quando a de dentro não trouxe definição própria — a que traz
                    // é a que se reiniciou, e herdar a de fora desfaria o reinício.
                    var fromParent = inherited is { } outer
                                     && definition is null
                                     && string.Equals(outer.Kind, node.Type, StringComparison.Ordinal)
                        ? (int?)outer.NumberingId
                        : null;

                    // O `numId` do arquivo quando o nó o trouxe; o da lista de
                    // fora quando esta é aninhada; e uma definição nova quando a
                    // lista nasceu aqui dentro — ou quando o `numId` que veio no
                    // modelo é de outro documento e este não o define, o que
                    // acontece em lista colada. O que não pode é sair zero, que no
                    // formato quer dizer "sem numeração": era assim que uma lista
                    // criada no editor voltava como parágrafo comum.
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

                        // O item dá a identidade; o parágrafo de dentro dá o
                        // conteúdo.
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
                    // Sem citação no OOXML: vira recuo, que é o que o Word faz.
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
    /// A largura da coluna de texto, em pixels do CSS.
    /// </summary>
    /// <remarks>
    /// É o teto de uma imagem que chega sem medida — a que a pessoa insere pela
    /// barra de ferramentas. Sem ele uma captura de tela de 1920 px entrava com
    /// meio metro de largura e o Word a desenhava por cima das duas margens.
    /// </remarks>
    private static int UsableWidthPx(PageSetupDto page)
    {
        var (shortSide, longSide) = PageReader.MillimetersOfPaper(page.Size);

        var across = string.Equals(page.Orientation, "landscape", StringComparison.Ordinal)
            ? longSide
            : shortSide;

        var millimeters = across - page.Margins.Left - page.Margins.Right;
        return millimeters > 10 ? (int)Math.Round(millimeters / 25.4 * 96) : ImageWriter.DefaultWidthPx;
    }

    internal static string? OidOf(Node node)
    {
        if (node.Attrs is null || !node.Attrs.TryGetValue("oid", out var value) || value is null) return null;
        return value.GetValueKind() == System.Text.Json.JsonValueKind.String ? value.GetValue<string>() : null;
    }

    /// <summary>
    /// O que havia dentro de um bloco editado e não sabemos regenerar.
    /// </summary>
    /// <remarks>
    /// Esta é a **perda de verdade**, e ela é detectada por comparação em vez
    /// de por palpite: o aviso pode dizer "você editou um parágrafo que tinha
    /// um comentário ancorado", em vez de um alerta genérico na abertura que o
    /// usuário aprende a ignorar.
    /// </remarks>
    /// <param name="beforeComments">
    /// O rascunho é de antes dos comentários (M10): os nós não trazem a âncora, e
    /// reescrever o parágrafo a perde. Depois dele a âncora volta pelos nós — só a
    /// referência que dividia o run com texto não tem como voltar.
    /// </param>
    /// <param name="beforeRevisions">
    /// O rascunho é de antes das revisões (M10): os nós não as trazem, e reescrever
    /// o parágrafo perde o `w:ins` e o `w:del`. Depois dele elas voltam pelas
    /// marcas — só a de formatação do trecho e a movimentação não voltam inteiras.
    /// </param>
    /// <param name="beforeNotes">
    /// O rascunho é de antes das notas (M11): os nós não trazem a referência, e
    /// reescrever o parágrafo a perde. Depois dele ela volta pelo `noteRef`.
    /// </param>
    private static void NoteWhatWasInside(
        Block block, Inventory inventory, bool beforeComments, bool beforeRevisions, bool beforeNotes)
    {
        var original = block.Source;

        if (beforeComments
                ? original.Descendants<CommentRangeStart>().Any() || original.Descendants<CommentReference>().Any()
                : original.Descendants<CommentReference>().Any(reference =>
                    reference.Parent is Run run && BodyReader.ReferenceOnly(run) is null))
        {
            inventory.NoteLoss("comentário ancorado num parágrafo que você editou");
        }

        // No rascunho antigo nenhuma revisão vem pelas marcas: a movimentação e a
        // de formatação também se perdem, e não só o `w:ins` e o `w:del`.
        if (beforeRevisions &&
            (original.Descendants<InsertedRun>().Any() || original.Descendants<DeletedRun>().Any() ||
             original.Descendants<MoveFromRun>().Any() || original.Descendants<MoveToRun>().Any() ||
             original.Descendants<MoveFromRangeStart>().Any() || original.Descendants<MoveToRangeStart>().Any() ||
             original.Descendants<RunPropertiesChange>().Any(change => change.Parent?.Parent is Run)))
        {
            inventory.NoteLoss("marcas de revisão num parágrafo que você editou");
        }

        // A imagem dentro de uma revisão não leva a marca (só o texto leva): o
        // parágrafo reescrito a devolve como conteúdo aceito.
        if (!beforeRevisions &&
            original.Descendants().Any(element =>
                element is InsertedRun or DeletedRun or MoveFromRun or MoveToRun &&
                (element.Descendants<Drawing>().Any() || element.Descendants<Picture>().Any() ||
                 element.Descendants<EmbeddedObject>().Any())))
        {
            inventory.NoteLoss("imagem dentro de uma revisão num parágrafo que você editou (ficou como aceita)");
        }

        // A formatação de antes da revisão mora no `w:rPr` do run, que a gravação
        // refaz a partir das marcas. A do parágrafo (`w:pPrChange`) e a da marca
        // de parágrafo voltam com o `w:pPr` original.
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

        // O campo que virou nó volta ao arquivo como campo (ver BodyReader.ReadField);
        // perde-se só o que o leitor mostrou pelo resultado.
        if (block.UnrepresentedField)
        {
            inventory.NoteLoss("campo calculado num parágrafo que você editou");
        }

        // Objeto ancorado não entra mais aqui: ele é copiado do original para o
        // parágrafo reescrito. Sobra o que continua sem volta — o VML antigo
        // (`w:pict`) e o desenho preso a um run que também traz texto, que
        // copiado traria a frase junto.
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
            // O papel do arquivo continua sendo o que o modelo diz — só a folha
            // girou, ou a margem mudou. Então as medidas dele ficam: um A5 não
            // tem por que virar A4 por causa de uma mudança de margem.
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

    /// <summary>
    /// As medidas do arquivo, quando ainda correspondem ao papel que o modelo
    /// nomeia.
    /// </summary>
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

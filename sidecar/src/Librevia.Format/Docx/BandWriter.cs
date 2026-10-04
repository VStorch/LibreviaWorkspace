using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O texto digitado na faixa volta a **um <c>w:t</c> por peça editada**: moldura,
/// tabela, logotipo e campo <c>PAGE</c> seguem byte a byte. Comparar a faixa inteira
/// seria regeneração disfarçada. Parte sem texto mudado nem entra na lista de
/// graváveis.
/// </summary>
internal static class BandWriter
{
    /// <summary>A mesma frase do main: é a mesma perda.</summary>
    private const string ForeignBands = "cabeçalho e rodapé do arquivo .docx de origem";

    /// <summary>Devolve os caminhos que mudaram.</summary>
    /// <param name="sections">Duas seções que apontam a mesma parte trazem o mesmo texto (PageReader.ReadAll).</param>
    internal static HashSet<string> Apply(
        MainDocumentPart part,
        IReadOnlyList<PageSetupDto?> sections,
        Inventory inventory,
        IReadOnlyDictionary<string, string>? aliases = null)
    {
        var touched = new HashSet<string>(StringComparer.Ordinal);
        var bands = sections.OfType<PageSetupDto>().SelectMany(BandsOf).ToList();
        if (bands.Count == 0) return touched;

        // Endereço → texto; a parte é uma só no arquivo.
        var wanted = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var band in bands)
        {
            foreach (var piece in PiecesOf(band))
            {
                if (Translate(piece.Pid, Relationship, aliases, inventory) is not { } pid) continue;
                Remember(wanted, piece with { Pid = pid });
            }
        }

        // O mesmo para as caixas, onde mora o título.
        var boxes = new Dictionary<string, List<Node>>(StringComparer.Ordinal);
        foreach (var band in bands)
        {
            foreach (var float_ in band.Floats ?? [])
            {
                if (float_.BoxId is null || float_.Kind != "text") continue;
                if (Translate(float_.BoxId, BoxRelationship, aliases, inventory) is not { } bid) continue;
                boxes[bid] = float_.Content ?? [];
            }
        }

        if (wanted.Count == 0 && boxes.Count == 0) return touched;

        var fonts = new FontTable(part);
        var relationships = wanted.Keys.Select(Relationship)
            .Concat(boxes.Keys.Select(BoxRelationship))
            .Where(id => id.Length > 0)
            .Distinct(StringComparer.Ordinal);

        foreach (var relationship in relationships)
        {
            // Sem a relação (o `.sdoc` reaberto num pacote mínimo), não há onde escrever: dito, e não calado.
            if (BandNav.PartOf(part, relationship) is not { } target)
            {
                inventory.NoteLoss(ForeignBands);
                continue;
            }

            // A mesma tabela de fontes do leitor: outra fusão de runs daria outra numeração de peças.
            var changed = ApplyTo(
                target.Root,
                wanted.Where(entry => Relationship(entry.Key) == relationship),
                inventory,
                fonts);

            changed |= ApplyBoxes(
                target.Root,
                boxes.Where(entry => BoxRelationship(entry.Key) == relationship),
                part,
                inventory,
                fonts);

            if (!changed) continue;

            target.Root.Save();
            touched.Add(BandNav.PathOf(target.Owner));
        }

        return touched;
    }

    /// <summary>Só as caixas cujo texto mudou, inteiras; as outras seguem com moldura e giro.</summary>
    private static bool ApplyBoxes(
        OpenXmlPartRootElement root,
        IEnumerable<KeyValuePair<string, List<Node>>> wanted,
        MainDocumentPart part,
        Inventory inventory,
        FontTable fonts)
    {
        var boxes = BandNav.BoxesOf(root);
        var writer = new ParagraphWriter(part, inventory);
        var touched = false;

        foreach (var (address, content) in wanted.Select(entry => (entry.Key, entry.Value)))
        {
            if (BandNav.ParseBox(address) is not { } at) continue;
            if (at.Box >= boxes.Count) continue;

            var box = boxes[at.Box];

            // Contra o texto da tela, com `{n}`: contra o XML a caixa mudaria sempre.
            if (HeaderReader.BoxTextOf(box, inventory, fonts) == PlainTextOf(content)) continue;

            // Caixa com campo não é reescrita: um cabeçalho que deixa de contar páginas é pior.
            if (HasField(box))
            {
                inventory.NoteLoss("texto de uma caixa de cabeçalho com campo calculado");
                continue;
            }

            // Reescrita, a caixa perde o marcador: avisa-se.
            if (box.Descendants<BookmarkStart>().Any())
            {
                inventory.NoteLoss("marcador numa caixa de cabeçalho que você editou");
            }

            if (box.Descendants<InsertedRun>().Any() || box.Descendants<DeletedRun>().Any() ||
                box.Descendants<MoveFromRun>().Any() || box.Descendants<MoveToRun>().Any())
            {
                inventory.NoteLoss("revisões numa caixa de cabeçalho que você editou");
            }

            box.RemoveAllChildren();
            foreach (var block in content)
            {
                foreach (var element in writer.Write(block)) box.AppendChild(element);
            }

            // `w:txbxContent` vazio invalida o documento.
            if (!box.HasChildren) box.AppendChild(new Paragraph());
            touched = true;
        }

        if (touched) Mirror(root);
        return touched;
    }

    private static bool HasField(TextBoxContent box) =>
        box.Descendants<FieldChar>().Any()
        || box.Descendants<FieldCode>().Any()
        || box.Descendants<SimpleField>().Any();

    private static string PlainTextOf(List<Node> content) =>
        string.Join("\n", content.Select(node => string.Concat(Texts(node))));

    private static IEnumerable<string> Texts(Node node)
    {
        if (node.Text is { } text) yield return text;
        foreach (var child in node.Content ?? [])
        {
            foreach (var value in Texts(child)) yield return value;
        }
    }

    private static string BoxRelationship(string address) =>
        BandNav.ParseBox(address) is { } parsed ? parsed.RelationshipId : string.Empty;

    private static IEnumerable<BandDto> BandsOf(PageSetupDto page)
    {
        foreach (var band in new[]
                 {
                     page.Header, page.Footer,
                     page.FirstHeader, page.FirstFooter,
                     page.EvenHeader, page.EvenFooter,
                 })
        {
            if (band is not null) yield return band;
        }
    }

    private static IEnumerable<PieceDto> PiecesOf(BandDto band)
    {
        foreach (var piece in band.Left) yield return piece;
        foreach (var piece in band.Center) yield return piece;
        foreach (var piece in band.Right) yield return piece;

        foreach (var row in band.Rows ?? [])
        {
            foreach (var cell in row.Cells)
            {
                foreach (var piece in cell.Pieces) yield return piece;
            }
        }
    }

    private static void Remember(Dictionary<string, string> wanted, PieceDto piece)
    {
        if (piece.Pid is null || piece.Kind != PieceDto.KindText) return;
        wanted[piece.Pid] = piece.Text ?? string.Empty;
    }

    /// <summary>A faixa desvinculada (<c>s2~rId5:0:1</c>) com a relação que a gravação criou (SectionWriter.ApplyBands).</summary>
    private static string? Translate(
        string? address,
        Func<string, string> relationshipOf,
        IReadOnlyDictionary<string, string>? aliases,
        Inventory inventory)
    {
        if (address is null) return null;
        var relationship = relationshipOf(address);
        if (!relationship.Contains(SectionWriter.UnlinkedSeparator, StringComparison.Ordinal)) return address;
        if (aliases is not null && aliases.TryGetValue(relationship, out var fresh) && fresh.Length > 0)
        {
            return fresh + address[relationship.Length..];
        }

        inventory.NoteLoss("cabeçalho ou rodapé desvinculado da seção anterior");
        return null;
    }

    private static string Relationship(string address) =>
        BandNav.Parse(address) is { } parsed ? parsed.RelationshipId : string.Empty;

    /// <summary>Devolve se algo mudou mesmo.</summary>
    private static bool ApplyTo(
        OpenXmlPartRootElement root,
        IEnumerable<KeyValuePair<string, string>> wanted,
        Inventory inventory,
        FontTable fonts)
    {
        var paragraphs = BandNav.ParagraphsOf(root);
        var cache = new Dictionary<int, List<HeaderReader.TracedPiece>>();
        var touched = false;

        foreach (var (address, text) in wanted.Select(entry => (entry.Key, entry.Value)))
        {
            if (BandNav.Parse(address) is not { } at) continue;
            if (at.Paragraph >= paragraphs.Count) continue;

            if (!cache.TryGetValue(at.Paragraph, out var pieces))
            {
                pieces = HeaderReader.TracedRuns(paragraphs[at.Paragraph], inventory, fonts);
                cache[at.Paragraph] = pieces;
            }

            if (at.Piece >= pieces.Count) continue;

            var piece = pieces[at.Piece];
            if (piece.Source.Count == 0) continue;
            if (piece.Piece.Text == text) continue;

            Rewrite(piece.Source, text, piece.Piece.Literal);
            touched = true;
        }

        if (!touched) return false;

        Mirror(root);
        return true;
    }

    /// <summary>O ramo de reserva em VML repete a caixa: escrever num só deixaria o arquivo dizendo duas coisas.</summary>
    private static void Mirror(OpenXmlPartRootElement root)
    {
        foreach (var alternate in root.Descendants<AlternateContent>().ToList())
        {
            TextBoxNav.MirrorFallback(alternate);
        }
    }

    /// <summary>
    /// O texto no primeiro <c>w:t</c> da peça, e os demais esvaziam: a contagem de runs,
    /// e com ela o endereço, fica estável. <c>xml:space="preserve"</c> sempre.
    /// </summary>
    private static void Rewrite(List<Text> source, string text, bool literal)
    {
        var segments = literal ? [new FieldTokens.Segment(null, text)] : FieldTokens.Split(text);
        source[0].Text = segments[0].Text;
        source[0].Space = SpaceProcessingModeValues.Preserve;

        for (var index = 1; index < source.Count; index++) source[index].Text = string.Empty;

        // `{n}` e `{total}` viram campos logo depois do run da peça, com a mesma formatação.
        if (segments.Count == 1 || source[0].Parent is not Run anchor) return;

        OpenXmlElement last = anchor;
        foreach (var segment in segments.Skip(1))
        {
            if (segment.Field is { } instruction)
            {
                var field = new SimpleField(RunLike(anchor, "1")) { Instruction = instruction };
                last = last.InsertAfterSelf(field);
            }

            if (segment.Text.Length > 0) last = last.InsertAfterSelf(RunLike(anchor, segment.Text));
        }
    }

    private static Run RunLike(Run model, string text)
    {
        var run = new Run();
        if (model.RunProperties is { } properties) run.AppendChild(properties.CloneNode(true));
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }
}

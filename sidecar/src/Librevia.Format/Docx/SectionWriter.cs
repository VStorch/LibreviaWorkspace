using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// As seções antes da última. No modelo o parágrafo que encerra a seção leva
/// <c>sectionBreak</c>, e a configuração mora em <c>sections</c>. O preservado volta com o
/// <c>w:sectPr</c> dele, regravado só no que difere; o reescrito o recebe de volta só
/// se ainda leva a marca; a marca nova ganha uma cópia do <c>w:sectPr</c> da seção que
/// partiu, com as referências de faixa.
/// </summary>
internal static class SectionWriter
{
    /// <param name="Existing">Nulo na marca nova.</param>
    internal sealed record Break(Paragraph Holder, SectionProperties? Existing, string Id);

    /// <summary>Ou nula.</summary>
    /// <param name="node">Para o item de lista, o parágrafo de dentro.</param>
    /// <param name="paragraphs">Os <c>w:p</c> que o bloco produziu.</param>
    /// <param name="kept">O bloco voltou byte a byte.</param>
    /// <param name="used">O parágrafo partido ou colado leva a marca repetida.</param>
    /// <param name="known">Marca de id que o modelo não configura não é quebra, e o <c>w:sectPr</c> dela sai.</param>
    public static Break? Mark(
        Node node,
        List<Paragraph> paragraphs,
        bool kept,
        HashSet<string> used,
        IReadOnlySet<string> known,
        Inventory inventory)
    {
        if (paragraphs.Count == 0) return null;

        var id = node.Type is "paragraph" or "heading" ? Attr.String(node, "sectionBreak") : null;
        if (id is not null && (!known.Contains(id) || !used.Add(id)))
        {
            // Sem seção no modelo, o `w:sectPr` sai, e se diz: a tela não a mostra.
            if (paragraphs.Any(p => p.ParagraphProperties?.SectionProperties is not null))
            {
                inventory.NoteLoss("quebra de seção sem configuração no documento (a seção foi unida à seguinte)");
            }

            id = null;
            kept = false;
        }

        if (kept)
        {
            // A marca não entra na impressão digital; sem marca, a quebra saiu, e o trecho passa à seção de baixo.
            var holder = paragraphs.FirstOrDefault(p => p.ParagraphProperties?.SectionProperties is not null);
            if (id is null)
            {
                holder?.ParagraphProperties?.SectionProperties?.Remove();
                return null;
            }

            return new Break(holder ?? paragraphs[^1], holder?.ParagraphProperties?.SectionProperties, id);
        }

        // O reescrito traz o `w:sectPr` pela cópia do `w:pPr`: volta ao último pedaço só se a marca continua.
        SectionProperties? carried = null;
        foreach (var paragraph in paragraphs)
        {
            if (paragraph.ParagraphProperties?.SectionProperties is not { } section) continue;
            carried ??= section;
            section.Remove();
        }

        if (id is null) return null;

        var last = paragraphs[^1];
        if (carried is not null) (last.ParagraphProperties ??= new ParagraphProperties()).SectionProperties = carried;
        return new Break(last, carried, id);
    }

    /// <returns>Os endereços de faixa desvinculada (<c>id~rIdN</c>) e a relação da parte nova de cada um.</returns>
    public static Dictionary<string, string> Apply(
        MainDocumentPart part,
        List<Break> breaks,
        SectionProperties last,
        DocumentModelDto model,
        Inventory inventory,
        HashSet<string> touched)
    {
        var aliases = new Dictionary<string, string>(StringComparer.Ordinal);

        var byId = (model.Sections ?? [])
            .Where(section => section.Id is not null)
            .GroupBy(section => section.Id!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        for (var index = 0; index < breaks.Count; index++)
        {
            var mark = breaks[index];
            var section = mark.Existing;

            if (section is null)
            {
                var template = breaks.Skip(index + 1).Select(next => next.Existing).FirstOrDefault(s => s is not null)
                               ?? last;
                section = (SectionProperties)template.CloneNode(true);
                (mark.Holder.ParagraphProperties ??= new ParagraphProperties()).SectionProperties = section;
                breaks[index] = mark with { Existing = section };
            }

            // A marca sem configuração fica com a da seção que copiou.
            if (!byId.TryGetValue(mark.Id, out var setup)) continue;

            if (!PageReader.Matches(section, setup)) DocxWriter.ApplyPageSetup(section, setup);
            ApplyStart(section, setup);
            ApplyColumns(section, setup);
            PageNumbering.Apply(part, section, setup, touched, inventory, documentWide: false);
        }

        var replaced = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < breaks.Count; index++)
        {
            if (breaks[index].Existing is { } section && byId.TryGetValue(breaks[index].Id, out var setup))
            {
                ApplyBands(part, section, setup, index, aliases, replaced);
            }
        }

        ApplyBands(part, last, model.Page, breaks.Count, aliases, replaced);
        DropOrphans(part, replaced);
        return aliases;
    }

    /// <summary>As partes de faixa que nenhuma seção aponta mais saem, senão cada gravação deixaria uma cópia.</summary>
    private static void DropOrphans(MainDocumentPart part, HashSet<string> replaced)
    {
        if (replaced.Count == 0) return;
        var referenced = part.Document?.Body?.Descendants<HeaderFooterReferenceType>()
            .Select(reference => reference.Id?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal) ?? [];
        foreach (var relationship in replaced.Where(id => !referenced.Contains(id)))
        {
            try
            {
                part.DeletePart(relationship);
            }
            catch (ArgumentOutOfRangeException)
            {
                // A relação já não existia.
            }
        }
    }

    /// <summary>Só o que difere e o painel conhece; o modelo sem larguras diferentes as iguala.</summary>
    public static void ApplyColumns(SectionProperties section, PageSetupDto page)
    {
        if (page.Columns is not { } wanted) return;
        var current = PageReader.ColumnsOf(section);
        var sameWidths = (wanted.WidthsMm ?? []).SequenceEqual(current.WidthsMm ?? []);
        if (wanted.Count == current.Count && wanted.SpaceMm == current.SpaceMm &&
            wanted.Separator == current.Separator && sameWidths)
        {
            return;
        }

        var columns = section.GetFirstChild<Columns>();
        if (columns is null)
        {
            columns = new Columns();
            if (!section.AddChild(columns, throwOnError: false)) return;
        }

        var count = Math.Clamp(wanted.Count, 1, 45);
        columns.ColumnCount = count > 1 ? (short)count : null;
        columns.Space = Attr.MmToTwips(wanted.SpaceMm).ToString(System.Globalization.CultureInfo.InvariantCulture);
        columns.Separator = wanted.Separator ? true : null;
        if (!sameWidths || wanted.WidthsMm is null)
        {
            columns.RemoveAllChildren<Column>();
            columns.EqualWidth = null;
        }
    }

    /// <summary>
    /// Desvincular copia a faixa da anterior, ainda apontando a parte de lá; o editor
    /// marca o endereço com a seção dona (<c>s2~rId5:0:1</c>), e aqui a marca vira parte nova.
    /// </summary>
    public const char UnlinkedSeparator = '~';

    private static string KeyOf(PageSetupDto setup) => setup.Id ?? "body";

    /// <summary>
    /// "Vincular ao anterior", tipo por tipo: faixa nula da segunda seção em diante é
    /// herança, e a referência sai; presente sem referência é desvinculada, e a parte
    /// anterior é copiada com imagens e links.
    /// </summary>
    private static void ApplyBands(
        MainDocumentPart part,
        SectionProperties section,
        PageSetupDto setup,
        int index,
        Dictionary<string, string> aliases,
        HashSet<string> replaced)
    {
        var key = KeyOf(setup);
        foreach (var (header, type, band) in new (bool, HeaderFooterValues, BandDto?)[]
                 {
                     (true, HeaderFooterValues.Default, setup.Header),
                     (true, HeaderFooterValues.First, setup.FirstHeader),
                     (true, HeaderFooterValues.Even, setup.EvenHeader),
                     (false, HeaderFooterValues.Default, setup.Footer),
                     (false, HeaderFooterValues.First, setup.FirstFooter),
                     (false, HeaderFooterValues.Even, setup.EvenFooter),
                 })
        {
            var references = (header
                    ? section.Elements<HeaderReference>().Cast<HeaderFooterReferenceType>()
                    : section.Elements<FooterReference>())
                .Where(reference => (reference.Type?.Value ?? HeaderFooterValues.Default) == type)
                .ToList();

            if (band is null)
            {
                if (index > 0) foreach (var reference in references) reference.Remove();
                continue;
            }

            var source = UnlinkedSource(band, key);
            if (source is null)
            {
                // A herdada que a seção passou a declarar (a de cima foi excluída) aponta a mesma parte.
                if (references.Count == 0 && SharedRelationship(band) is { } shared && PartExists(part, shared))
                {
                    AddReference(section, header, type, shared);
                }

                continue;
            }

            // Sempre uma parte nova, copiada da herdada: a que a seção já aponta pode ser a de antes.
            var relationship = source[(key.Length + 1)..];
            if (CloneBandPart(part, relationship, header) is not { } fresh) continue;
            foreach (var reference in references)
            {
                if (reference.Id?.Value is { } old) replaced.Add(old);
                reference.Remove();
            }

            AddReference(section, header, type, fresh);
            aliases[source] = fresh;
        }
    }

    private static void AddReference(SectionProperties section, bool header, HeaderFooterValues type, string id)
    {
        HeaderFooterReferenceType added = header
            ? new HeaderReference { Type = type, Id = id }
            : new FooterReference { Type = type, Id = id };
        // Na ordem do esquema: cabeçalhos antes dos rodapés, abrindo o `w:sectPr`.
        var after = section.ChildElements
            .Where(child => header ? child is HeaderReference : child is HeaderReference or FooterReference)
            .LastOrDefault();
        if (after is null) section.InsertAt(added, 0);
        else section.InsertAfter(added, after);
    }

    private static bool PartExists(MainDocumentPart part, string relationship) =>
        part.Parts.Any(pair => pair.RelationshipId == relationship);

    private static string? SharedRelationship(BandDto band)
    {
        var address = AddressesOf(band).FirstOrDefault(value => value is not null && !value.Contains(UnlinkedSeparator));
        if (address is null) return null;
        var end = address.IndexOfAny([':', '#']);
        return end <= 0 ? null : address[..end];
    }

    private static IEnumerable<string?> AddressesOf(BandDto band) =>
        band.Left.Concat(band.Center).Concat(band.Right)
            .Concat((band.Rows ?? []).SelectMany(row => row.Cells).SelectMany(cell => cell.Pieces))
            .Select(piece => piece.Pid)
            .Concat((band.Floats ?? []).Select(item => item.BoxId));

    private static string? UnlinkedSource(BandDto band, string key)
    {
        var prefix = key + UnlinkedSeparator;
        foreach (var address in AddressesOf(band))
        {
            if (address is null || !address.StartsWith(prefix, StringComparison.Ordinal)) continue;
            var end = address.IndexOfAny([':', '#'], prefix.Length);
            return end < 0 ? address : address[..end];
        }

        return null;
    }

    /// <summary>Com as mesmas imagens e links.</summary>
    private static string? CloneBandPart(MainDocumentPart part, string relationship, bool header)
    {
        if (string.IsNullOrEmpty(relationship)) return null;
        OpenXmlPart? source;
        try
        {
            source = part.GetPartById(relationship);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }

        OpenXmlPart target = header ? part.AddNewPart<HeaderPart>() : part.AddNewPart<FooterPart>();
        using (var stream = source.GetStream()) target.FeedData(stream);
        foreach (var child in source.Parts) target.AddPart(child.OpenXmlPart, child.RelationshipId);
        foreach (var external in source.ExternalRelationships)
        {
            target.AddExternalRelationship(external.RelationshipType, external.Uri, external.Id);
        }

        foreach (var link in source.HyperlinkRelationships)
        {
            target.AddHyperlinkRelationship(link.Uri, link.IsExternal, link.Id);
        }

        return part.GetIdOfPart(target);
    }

    /// <summary>Só quando difere do arquivo; ausente é "não mexa".</summary>
    public static void ApplyStart(SectionProperties section, PageSetupDto page)
    {
        if (page.Start is not { } start || !PageReader.SectionStarts.Contains(start)) return;
        if (start == PageReader.StartOf(section)) return;

        section.RemoveAllChildren<SectionType>();
        // "Próxima página" é o padrão: sem elemento, como o Word grava.
        if (start == "nextPage") return;

        section.AddChild(new SectionType { Val = ValueOf(start) }, throwOnError: false);
    }

    private static SectionMarkValues ValueOf(string start) => start switch
    {
        "continuous" => SectionMarkValues.Continuous,
        "evenPage" => SectionMarkValues.EvenPage,
        "oddPage" => SectionMarkValues.OddPage,
        "nextColumn" => SectionMarkValues.NextColumn,
        _ => SectionMarkValues.NextPage,
    };
}

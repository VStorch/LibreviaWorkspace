using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// As seções antes da última, levadas de volta ao arquivo.
/// </summary>
/// <remarks>
/// No OOXML a seção termina num `w:sectPr` guardado no `w:pPr` do parágrafo que
/// a encerra; no modelo esse parágrafo leva `sectionBreak` com o id da seção, e a
/// configuração dela mora em `sections`, fora dos nós. As duas pontas se
/// encontram aqui, depois que o corpo foi montado:
///
/// - o parágrafo preservado volta com o `w:sectPr` dele, e só o que o modelo diz
///   diferente é regravado — papel, margens, começo, numeração —, pelas mesmas
///   comparações da última seção;
/// - o parágrafo reescrito perde o `w:sectPr` que ParagraphFormat copiou do
///   original, e o recebe de volta só se ainda leva a marca: apagar a marca é
///   apagar a quebra, e o trecho passa à seção de baixo, como no Word;
/// - a marca nova (Inserir → Quebra de seção) ganha uma cópia do `w:sectPr` da
///   seção que ela partiu — a seguinte, no arquivo —, com as referências de
///   faixa junto: as duas metades continuam mostrando o mesmo cabeçalho, que é o
///   que o modelo também faz ao copiar a seção.
/// </remarks>
internal static class SectionWriter
{
    /// <summary>Um parágrafo que encerra seção, já no corpo que vai ser gravado.</summary>
    /// <param name="Existing">O `w:sectPr` que ele já tinha; nulo na marca nova.</param>
    internal sealed record Break(Paragraph Holder, SectionProperties? Existing, string Id);

    /// <summary>
    /// A marca de seção do bloco recém-gravado, ou nula.
    /// </summary>
    /// <param name="node">O nó do parágrafo (para o item de lista, o de dentro).</param>
    /// <param name="paragraphs">Os `w:p` que o bloco produziu.</param>
    /// <param name="kept">O bloco voltou byte a byte.</param>
    /// <param name="used">Os ids já vistos: o parágrafo partido ou colado leva a marca repetida.</param>
    /// <param name="known">
    /// Os ids que o modelo configura. Marca de outro id — parágrafo colado de
    /// outro documento, marca que voltou por um desfazer sem a seção — não é
    /// quebra: a tela também não a trata como uma, e o `w:sectPr` dela sai.
    /// </param>
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
            // Preservado com `w:sectPr` e sem seção no modelo: guardá-lo faria o
            // arquivo ter uma seção que a tela não mostra, e deslocaria a conta
            // das seguintes. Sai, e se diz.
            if (paragraphs.Any(p => p.ParagraphProperties?.SectionProperties is not null))
            {
                inventory.NoteLoss("quebra de seção sem configuração no documento (a seção foi unida à seguinte)");
            }

            id = null;
            kept = false;
        }

        if (kept)
        {
            // O preservado volta com o `w:sectPr` dele — a marca não entra na
            // impressão digital, e o parágrafo cuja seção foi renomeada ou cuja
            // quebra foi excluída continua preservado. Sem marca, a quebra saiu:
            // o `w:sectPr` sai também, e o trecho passa à seção de baixo.
            var holder = paragraphs.FirstOrDefault(p => p.ParagraphProperties?.SectionProperties is not null);
            if (id is null)
            {
                holder?.ParagraphProperties?.SectionProperties?.Remove();
                return null;
            }

            return new Break(holder ?? paragraphs[^1], holder?.ParagraphProperties?.SectionProperties, id);
        }

        // O reescrito traz o `w:sectPr` do original pela cópia do `w:pPr`: sai
        // de todos os pedaços, e volta ao último só se a marca continua.
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

    /// <summary>
    /// Dá a cada marca o `w:sectPr` que o modelo descreve.
    /// </summary>
    /// <returns>
    /// Os endereços de faixa desvinculada (`id~rIdN`) e a relação da parte nova
    /// que cada um passou a ter — ver <see cref="ApplyBands"/>.
    /// </returns>
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
                // A seção que a marca nova partiu é a seguinte que já existia.
                var template = breaks.Skip(index + 1).Select(next => next.Existing).FirstOrDefault(s => s is not null)
                               ?? last;
                section = (SectionProperties)template.CloneNode(true);
                (mark.Holder.ParagraphProperties ??= new ParagraphProperties()).SectionProperties = section;
                breaks[index] = mark with { Existing = section };
            }

            // A marca sem configuração no modelo (o id veio de outro documento,
            // num parágrafo colado) fica com a da seção que ela copiou.
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

    /// <summary>
    /// As partes de faixa cuja referência a gravação trocou, e que nenhuma seção
    /// aponta mais: saem, senão cada gravação deixaria uma cópia esquecida.
    /// </summary>
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
                // A relação já não existia: nada a apagar.
            }
        }
    }

    /// <summary>
    /// `w:cols`: só quando o modelo diz algo diferente do arquivo, e só o que o
    /// painel conhece. Larguras diferentes voltam como estavam enquanto o modelo
    /// as trouxer; o modelo sem elas (a pessoa mudou as colunas) as iguala.
    /// </summary>
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

    /// <summary>O prefixo que o editor põe no endereço da faixa desvinculada.</summary>
    /// <remarks>
    /// Desvincular copia a faixa da seção anterior, e a cópia ainda aponta a parte
    /// de lá. O editor marca cada endereço copiado com a seção dona
    /// (`s2~rId5:0:1`; a do corpo é `body`), e é aqui que a marca vira uma parte
    /// nova: editar a cópia não pode mudar o cabeçalho das outras seções.
    /// </remarks>
    public const char UnlinkedSeparator = '~';

    /// <summary>A chave da seção no endereço desvinculado.</summary>
    private static string KeyOf(PageSetupDto setup) => setup.Id ?? "body";

    /// <summary>
    /// "Vincular ao anterior", tipo por tipo.
    /// </summary>
    /// <remarks>
    /// Faixa nula numa seção que não é a primeira é herança: a referência que o
    /// arquivo tinha sai. Faixa presente sem referência é a desvinculada: a parte
    /// da seção anterior é copiada, com imagens e links, e a seção passa a
    /// apontá-la. Na primeira seção, nula é "sem faixa" desde a leitura (vazia
    /// também vira nula lá), e nada se remove.
    /// </remarks>
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
                // Faixa própria sem referência e sem marca de desvinculada: é a
                // herdada que a seção passou a declarar (a de cima foi excluída).
                // Aponta a mesma parte — é o mesmo cabeçalho.
                if (references.Count == 0 && SharedRelationship(band) is { } shared && PartExists(part, shared))
                {
                    AddReference(section, header, type, shared);
                }

                continue;
            }

            // A desvinculada ganha sempre uma parte nova, copiada da que ela
            // herdava — ainda que a seção já aponte uma: essa pode ser a própria
            // de antes, e as peças que o modelo traz são as da cópia.
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
        // As referências abrem o `w:sectPr`, cabeçalhos antes dos rodapés.
        var after = section.ChildElements
            .Where(child => header ? child is HeaderReference : child is HeaderReference or FooterReference)
            .LastOrDefault();
        if (after is null) section.InsertAt(added, 0);
        else section.InsertAfter(added, after);
    }

    private static bool PartExists(MainDocumentPart part, string relationship) =>
        part.Parts.Any(pair => pair.RelationshipId == relationship);

    /// <summary>A relação das peças de uma faixa sem marca de desvinculada (`rIdN:0:1`).</summary>
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

    /// <summary>A relação marcada (`chave~rIdN`) da faixa desvinculada, se a faixa é uma.</summary>
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

    /// <summary>
    /// Uma cópia da parte de cabeçalho ou rodapé, com as mesmas imagens e links.
    /// </summary>
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

    /// <summary>
    /// `w:type`: como a seção começa. Só quando o modelo diz algo diferente do
    /// arquivo — ausente (rascunho de antes) é "não mexa".
    /// </summary>
    public static void ApplyStart(SectionProperties section, PageSetupDto page)
    {
        if (page.Start is not { } start || !PageReader.SectionStarts.Contains(start)) return;
        if (start == PageReader.StartOf(section)) return;

        section.RemoveAllChildren<SectionType>();
        // "Próxima página" é o padrão da especificação: sem elemento, como o
        // Word grava.
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

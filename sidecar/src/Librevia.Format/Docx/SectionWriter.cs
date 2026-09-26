using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// As seções antes da última, levadas de volta ao arquivo (M9).
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
    public static Break? Mark(Node node, List<Paragraph> paragraphs, bool kept, HashSet<string> used)
    {
        if (paragraphs.Count == 0) return null;

        var id = node.Type is "paragraph" or "heading" ? Attr.String(node, "sectionBreak") : null;
        if (id is not null && !used.Add(id)) id = null;

        if (kept)
        {
            // O preservado com marca volta com o `w:sectPr` dele; sem marca, a
            // leitura de referência também não a deu, e ele não tem `w:sectPr`.
            var holder = paragraphs.FirstOrDefault(p => p.ParagraphProperties?.SectionProperties is not null);
            return id is null
                ? null
                : new Break(holder ?? paragraphs[^1], holder?.ParagraphProperties?.SectionProperties, id);
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
    public static void Apply(
        MainDocumentPart part,
        List<Break> breaks,
        SectionProperties last,
        DocumentModelDto model,
        Inventory inventory,
        HashSet<string> touched)
    {
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
            PageNumbering.Apply(part, section, setup, touched, inventory, documentWide: false);
        }
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

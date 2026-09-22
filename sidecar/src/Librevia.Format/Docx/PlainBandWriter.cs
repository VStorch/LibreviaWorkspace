using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// O cabeçalho e o rodapé de texto simples, levados ao arquivo.
/// </summary>
/// <remarks>
/// O documento novo não tem faixa preservada: tem uma linha de texto em
/// "Configurar página", com `{n}` e `{total}` no lugar dos números, que o PDF
/// desenha centralizada em Calibri 9 pt (`src/services/pdf/page-setup.ts`). Sem
/// esta ponte, salvar o documento novo em DOCX apagava as duas linhas sem aviso.
///
/// A parte gerada aqui é reconhecida pelo id da relação, que é nosso. É isso que
/// permite regravá-la quando a pessoa muda o texto e apagá-la quando o texto
/// some, sem nunca pôr a mão no cabeçalho de um documento que veio de fora: esse
/// volta do leitor como faixa, e a faixa manda — como no PDF.
/// </remarks>
internal static class PlainBandWriter
{
    private const string HeaderId = "LibreviaHeader";
    private const string FooterId = "LibreviaFooter";

    public static void Apply(
        MainDocumentPart part,
        SectionProperties section,
        PageSetupDto page,
        Inventory? inventory = null,
        HashSet<string>? touched = null)
    {
        var added = false;

        // Faixa com conteúdo manda, como no PDF — e é o caso do DOCX reaberto,
        // cujo cabeçalho (inclusive o que nasceu aqui) volta do leitor como faixa.
        // Aí a linha de texto nem é olhada, e muito menos apaga a parte.
        if (page.Header is not { IsEmpty: false })
        {
            added |= ApplyOne<HeaderReference, HeaderPart>(part, section, Meaningful(page.HeaderText), HeaderId,
                "cabeçalho", inventory, touched, content => new Header(content));
        }

        if (page.Footer is not { IsEmpty: false })
        {
            added |= ApplyOne<FooterReference, FooterPart>(part, section, Meaningful(page.FooterText), FooterId,
                "rodapé", inventory, touched, content => new Footer(content));
        }

        // Só quando entrou referência nova: reordenar as de um documento alheio
        // seria mexer no que ninguém pediu.
        if (added) Reorder(section);
    }

    private static string? Meaningful(string? text) => string.IsNullOrWhiteSpace(text) ? null : text;

    /// <returns>Se uma referência nova entrou no `w:sectPr`.</returns>
    private static bool ApplyOne<TReference, TPart>(
        MainDocumentPart part,
        SectionProperties section,
        string? text,
        string ownId,
        string label,
        Inventory? inventory,
        HashSet<string>? touched,
        Func<Paragraph, OpenXmlPartRootElement> root)
        where TReference : HeaderFooterReferenceType, new()
        where TPart : OpenXmlPart, IFixedContentTypePart
    {
        var existing = section.Elements<TReference>()
            .FirstOrDefault(reference => reference.Type is null || reference.Type.Value == HeaderFooterValues.Default);

        if (text is null)
        {
            // O texto sumiu e a parte é nossa: ela sai junto, senão o arquivo
            // continuaria imprimindo o que a tela já não mostra.
            if (existing?.Id?.Value == ownId)
            {
                existing.Remove();
                part.DeletePart(ownId);
            }

            return false;
        }

        if (existing is not null && existing.Id?.Value != ownId)
        {
            // O documento tem cabeçalho próprio que o leitor não soube mostrar.
            // Trocá-lo pela linha de texto seria perder o dele; ignorar a linha
            // em silêncio, perder a da pessoa. Fica o dele, e o aviso.
            inventory?.NoteLoss($"{label} de texto simples: o documento já tem um {label} próprio");
            return false;
        }

        var paragraph = Paragraph(text);

        if (existing is not null)
        {
            if (!part.TryGetPartById(ownId, out var found) || found is not TPart owned ||
                owned.RootElement is not { } current)
            {
                return false;
            }

            // Pela assinatura, e não pelo XML: o da parte lida traz declarações
            // de espaço de nomes que o parágrafo montado aqui não tem.
            if (Signature(current) == Signature(paragraph)) return false;

            current.RemoveAllChildren();
            current.AppendChild(paragraph);
            current.Save();
            touched?.Add(owned.Uri.OriginalString.TrimStart('/'));
            return false;
        }

        var created = part.AddNewPart<TPart>(ownId);
        using (var stream = created.GetStream(FileMode.Create)) root(paragraph).Save(stream);

        section.PrependChild(new TReference { Type = HeaderFooterValues.Default, Id = ownId });
        return true;
    }

    /// <summary>O texto e os campos, na ordem — o que a linha de texto simples diz.</summary>
    private static string Signature(OpenXmlElement element) => string.Concat(element.Descendants().Select(child =>
        child switch
        {
            Text text => text.Text,
            SimpleField field => $"{{{field.Instruction?.Value}}}",
            _ => string.Empty,
        }));

    /// <summary>
    /// A linha como o PDF a desenha: centralizada, Calibri 9 pt, cinza.
    /// </summary>
    private static Paragraph Paragraph(string text)
    {
        var paragraph = new Paragraph(new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "0" },
            new Justification { Val = JustificationValues.Center }));

        var rest = text;
        while (rest.Length > 0)
        {
            var page = rest.IndexOf("{n}", StringComparison.Ordinal);
            var total = rest.IndexOf("{total}", StringComparison.Ordinal);
            var next = new[] { page, total }.Where(at => at >= 0).DefaultIfEmpty(-1).Min();

            if (next < 0)
            {
                paragraph.AppendChild(Run(rest));
                break;
            }

            if (next > 0) paragraph.AppendChild(Run(rest[..next]));

            var isPage = next == page;
            paragraph.AppendChild(new SimpleField(Run("1")) { Instruction = isPage ? " PAGE " : " NUMPAGES " });
            rest = rest[(next + (isPage ? "{n}" : "{total}").Length)..];
        }

        return paragraph;
    }

    private static Run Run(string text) => new(
        new RunProperties(
            new RunFonts { Ascii = TemplateStyles.BandFont, HighAnsi = TemplateStyles.BandFont },
            new Color { Val = "444444" },
            new FontSize { Val = "18" }),
        new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    /// <summary>
    /// As referências abrem o `w:sectPr`, cabeçalhos antes dos rodapés.
    /// </summary>
    /// <remarks>
    /// O esquema exige que venham antes do papel; a ordem entre elas é do Word, que
    /// é quem mais estranha um arquivo diferente do que ele mesmo grava.
    /// </remarks>
    private static void Reorder(SectionProperties section)
    {
        var references = section.Elements<HeaderReference>().Cast<OpenXmlElement>()
            .Concat(section.Elements<FooterReference>())
            .ToList();

        foreach (var reference in references) reference.Remove();
        for (var index = references.Count - 1; index >= 0; index--) section.PrependChild(references[index]);
    }
}

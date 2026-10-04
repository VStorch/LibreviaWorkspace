using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// A linha de texto de "Configurar página" (com <c>{n}</c> e <c>{total}</c>) levada
/// ao DOCX. A parte gerada é reconhecida pelo id da relação, que é nosso; cabeçalho
/// de fora volta do leitor como faixa, e a faixa manda, como no PDF.
/// </summary>
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

        // A faixa com conteúdo manda, também a que nasceu aqui e voltou do leitor.
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

        // A ordem de um documento alheio só muda quando entrou referência nova.
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
            // A parte é nossa: sai com o texto, senão imprimiria o que a tela não mostra.
            if (existing?.Id?.Value == ownId)
            {
                existing.Remove();
                part.DeletePart(ownId);
            }

            return false;
        }

        if (existing is not null && existing.Id?.Value != ownId)
        {
            // Cabeçalho próprio que o leitor não mostra: fica o dele, com aviso.
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

            // A parte lida traz declarações de espaço de nomes que a montada não tem.
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

    private static string Signature(OpenXmlElement element) => string.Concat(element.Descendants().Select(child =>
        child switch
        {
            Text text => text.Text,
            SimpleField field => $"{{{field.Instruction?.Value}}}",
            _ => string.Empty,
        }));

    /// <summary>Como o PDF a desenha: centralizada, Calibri 9 pt, cinza.</summary>
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
    /// O esquema põe as referências antes do papel; cabeçalhos antes de rodapés é a
    /// ordem do Word.
    /// </summary>
    private static void Reorder(SectionProperties section)
    {
        var references = section.Elements<HeaderReference>().Cast<OpenXmlElement>()
            .Concat(section.Elements<FooterReference>())
            .ToList();

        foreach (var reference in references) reference.Remove();
        for (var index = references.Count - 1; index >= 0; index--) section.PrependChild(references[index]);
    }
}

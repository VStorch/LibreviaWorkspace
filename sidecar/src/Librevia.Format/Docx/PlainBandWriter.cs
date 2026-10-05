using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// The "Page setup" text line (with <c>{n}</c> and <c>{total}</c>) taken to the DOCX. The generated
/// part is recognized by the relationship id, which is ours; a header from outside comes back from
/// the reader as a band, and the band wins, as in the PDF.
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

        // A band with content wins, including one born here that came back from the reader.
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

        // Someone else's document only changes order when a new reference came in.
        if (added) Reorder(section);
    }

    private static string? Meaningful(string? text) => string.IsNullOrWhiteSpace(text) ? null : text;

    /// <returns>Whether a new reference entered the `w:sectPr`.</returns>
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
            // The part is ours: it goes with the text, or it would print what the screen does not
            // show.
            if (existing?.Id?.Value == ownId)
            {
                existing.Remove();
                part.DeletePart(ownId);
            }

            return false;
        }

        if (existing is not null && existing.Id?.Value != ownId)
        {
            // A header of its own that the reader does not show: its own stays, with a warning.
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

            // The read part carries namespace declarations the built one lacks.
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

    /// <summary>As the PDF draws it: centered, Calibri 9 pt, grey.</summary>
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
    /// The schema puts references before the paper; headers before footers is Word's order.
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

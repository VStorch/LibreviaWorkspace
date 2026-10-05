using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Librevia.Format.Docx;

/// <summary>
/// Sections before the last. In the model the paragraph closing a section carries
/// <c>sectionBreak</c>, and the setup lives in <c>sections</c>. A preserved paragraph goes back
/// with its <c>w:sectPr</c>, rewritten only where it differs; a rewritten one gets it back only if
/// it still carries the mark; a new mark gets a copy of the split section's <c>w:sectPr</c>, with
/// the band references.
/// </summary>
internal static class SectionWriter
{
    /// <param name="Existing">Null for a new mark.</param>
    internal sealed record Break(Paragraph Holder, SectionProperties? Existing, string Id);

    /// <summary>Or null.</summary>
    /// <param name="node">For a list item, the inner paragraph.</param>
    /// <param name="paragraphs">The <c>w:p</c>s the block produced.</param>
    /// <param name="kept">The block went back byte for byte.</param>
    /// <param name="used">A split or pasted paragraph carries the repeated mark.</param>
    /// <param name="known">A mark whose id the model does not configure is not a break, and its
    /// <c>w:sectPr</c> goes.</param>
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
            // Without a section in the model, the `w:sectPr` goes, and that is said: the screen
            // does not show it.
            if (paragraphs.Any(p => p.ParagraphProperties?.SectionProperties is not null))
            {
                inventory.NoteLoss("quebra de seção sem configuração no documento (a seção foi unida à seguinte)");
            }

            id = null;
            kept = false;
        }

        if (kept)
        {
            // The mark does not enter the fingerprint; without a mark, the break went away, and the
            // range passes to the section below.
            var holder = paragraphs.FirstOrDefault(p => p.ParagraphProperties?.SectionProperties is not null);
            if (id is null)
            {
                holder?.ParagraphProperties?.SectionProperties?.Remove();
                return null;
            }

            return new Break(holder ?? paragraphs[^1], holder?.ParagraphProperties?.SectionProperties, id);
        }

        // A rewritten paragraph brings the `w:sectPr` through the copied `w:pPr`: it goes back to
        // the last piece only if the mark remains.
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

    /// <returns>The unlinked band addresses (<c>id~rIdN</c>) and each one's new part
    /// relationship.</returns>
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

            // A mark without setup keeps that of the section it copied.
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
    /// Band parts no section points to anymore go, or every save would leave a copy.
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
                // The relationship no longer existed.
            }
        }
    }

    /// <summary>
    /// Only what differs and the panel knows; a model without unequal widths equalizes them.
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

    /// <summary>
    /// Unlinking copies the previous band, still pointing to its part; the editor tags the address
    /// with the owning section (<c>s2~rId5:0:1</c>), and here the tag becomes a new part.
    /// </summary>
    public const char UnlinkedSeparator = '~';

    private static string KeyOf(PageSetupDto setup) => setup.Id ?? "body";

    /// <summary>
    /// "Link to previous", kind by kind: a null band from the second section on is inheritance, and
    /// the reference goes; present without a reference means unlinked, and the previous part is
    /// copied with images and links.
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
                // An inherited band the section started declaring (the one above was deleted)
                // points to the same part.
                if (references.Count == 0 && SharedRelationship(band) is { } shared && PartExists(part, shared))
                {
                    AddReference(section, header, type, shared);
                }

                continue;
            }

            // Always a new part, copied from the inherited one: the one the section already points
            // to may be the previous one.
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
        // In schema order: headers before footers, opening the `w:sectPr`.
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

    /// <summary>With the same images and links.</summary>
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

    /// <summary>Only when it differs from the file; absent means "leave alone".</summary>
    public static void ApplyStart(SectionProperties section, PageSetupDto page)
    {
        if (page.Start is not { } start || !PageReader.SectionStarts.Contains(start)) return;
        if (start == PageReader.StartOf(section)) return;

        section.RemoveAllChildren<SectionType>();
        // "Next page" is the default: no element, as Word writes it.
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
